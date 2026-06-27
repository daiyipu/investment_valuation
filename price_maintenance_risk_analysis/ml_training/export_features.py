#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增特征导出脚本 - 从多个数据源汇总全部特征供ML训练

数据源:
  1. investment_valuation MySQL - 行情/估值/FCF/筛选/行业/定增参数
  2. investment_valuation MySQL - 财务评分/子场景/7个月涨跌幅
     (placement_evaluation + company_annual_scores，由 backfill_evaluations.py / batch 落库)
  3. fund_risk_control MySQL - 财务比率/资产负债表/利润表/现金流量表

用法:
    python ml_training/export_features.py [--output features.parquet]
    # 评分/标签脊柱完全来自 DB (placement_evaluation)，无需 Excel

输出:
    ml_training/data/features.parquet - 一行一只股票，所有特征平铺
"""

import sys
import os
import pymysql
import argparse
import json
import numpy as np
import pandas as pd


def _calc_features_from_series(prices, windows=[20, 60, 120, 250]):
    """从价格序列计算行情特征（用于报价日与market_data不匹配的样本）"""
    n = len(prices)
    if n < 20:
        return {}
    f = {}
    for w in windows:
        if n >= w:
            sub = np.array(prices[-w:], dtype=float)
            rets = np.diff(np.log(sub))
            f[f'波动率_{w}d'] = float(np.std(rets) * np.sqrt(250))
            f[f'年化收益_{w}d'] = float((sub[-1] / sub[0]) ** (250 / w) - 1) if sub[0] > 0 else None
            f[f'区间收益_{w}d'] = float(sub[-1] / sub[0] - 1) if sub[0] > 0 else None
            f[f'胜率_{w}d'] = float(np.mean(sub[1:] > sub[:-1]))
            f[f'MA{w}'] = float(np.mean(sub))
    f['当前价'] = float(prices[-1])
    f['数据天数'] = n
    if n >= 2:
        rets_all = np.diff(np.log(np.array(prices, dtype=float)))
        f['漂移率'] = float(np.mean(rets_all) * 250)
        f['波动率'] = float(np.std(rets_all) * np.sqrt(250))
    return f


FCF_COL_MAP = {'revenue': '营收', 'operate_profit': '营业利润', 'net_income': '净利润_fcf',
               'nopat': 'NOPAT', 'depreciation': '折旧', 'capex': '资本支出',
               'wc_change': '营运资金变动', 'fcf': 'FCF'}


def _pit_year(issue_date):
    """PIT 年: 月≥5 用上年(已披露), 否则前年(同 company_annual_scores 规则)。FCF/评分共用。"""
    s = str(issue_date)
    byr = int(s[:4]); bmo = int(s[4:6]) if len(s) >= 6 else 13
    return byr - 1 if bmo >= 5 else byr - 2


def _load_fcf_cache(codes):
    """批量取 historical_fcf(一条 IN 查询) → groupby 成 {code: df_sorted_by_year_desc}。"""
    codes = [str(c) for c in codes]
    if not codes:
        return {}
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    ph = ','.join(['%s'] * len(codes))
    eng = list(FCF_COL_MAP.keys())
    df = pd.read_sql(f"SELECT stock_code,year,{','.join(eng)} FROM historical_fcf "
                     f"WHERE stock_code IN ({ph})", conn, params=codes)
    conn.close()
    return {str(c): g.sort_values('year', ascending=False) for c, g in df.groupby('stock_code')}


_FCF_PIT_MEMO = {}   # {(code, pit_year): out_dict}  534400 key → ~85k 唯一组合


def _fcf_pit_cols(code, issue_date, fcf_cache):
    """一个样本的 FCF 基列(T/T1-4 + FCF年份 + 3年斜率), PIT year≤pit_year。

    忠实复刻原 load_db_features 段3 逻辑(保证定增输出不变), placement(load_db_features 内部)
    与回测(load_fcf_bulk)共用此函数 = FCF 装载单源。
    按 (code, pit_year) memoize + 向量化(arr 代替 iterrows): 全A 534400 key 从 8min+ → ~1min。
    """
    py = _pit_year(issue_date)
    mkey = (str(code), py)
    cached = _FCF_PIT_MEMO.get(mkey)
    if cached is not None:
        return cached
    grp = fcf_cache.get(str(code))
    out = {}
    if grp is not None and not grp.empty:
        pit = grp[grp['year'] <= py].head(5)
        n = len(pit)
        if n:
            arr = {c: pit[c].values for c in ['year'] + list(FCF_COL_MAP.keys())}
            for i in range(n):
                s = '_T' if i == 0 else f'_T{i}'
                for eng, cn in FCF_COL_MAP.items():
                    out[f'{cn}{s}'] = arr[eng][i]
                out[f'FCF年份{s}'] = arr['year'][i]
            if n >= 3:
                for col, nm in [('fcf', 'FCF'), ('revenue', '营收'), ('nopat', 'NOPAT')]:
                    vals = [arr[col][j] for j in range(3)]
                    vals = [v for v in vals if pd.notna(v)]
                    if len(vals) >= 2:
                        out[f'{nm}_3年斜率'] = np.polyfit(range(len(vals)), vals, 1)[0]
    _FCF_PIT_MEMO[mkey] = out
    return out


def load_fcf_bulk(sample_keys):
    """批量 FCF 基列(public): 一条 IN 查询 + 逐样本 PIT 选择(_fcf_pit_cols)。

    **FCF 装载单源**: placement(load_db_features 段3 内部调 _fcf_pit_cols) 与回测(build_backtest_panel)
    共用同一套 PIT 逻辑, 不再各编。
    """
    _FCF_PIT_MEMO.clear()   # 每次装入前清 memo(免跨 run 旧值)
    cache = _load_fcf_cache({str(c) for c, _ in sample_keys})
    return pd.DataFrame([_fcf_pit_cols(str(c), d, cache) for c, d in sample_keys])


def load_db_features(sample_keys):
    """从 investment_valuation MySQL 加载全部可关联特征

    Args:
        sample_keys: list of (stock_code, issue_date) 元组
    """
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4',
                           cursorclass=pymysql.cursors.DictCursor)
    cur = conn.cursor()

    mismatch_count = 0

    # 批量预取 FCF(groupby cache; 段3 复用 _fcf_pit_cols, 与回测 load_fcf_bulk 单源)
    fcf_cache = _load_fcf_cache({str(c) for c, _ in sample_keys})

    features = []
    for code, issue_date in sample_keys:
        f = {'股票代码': code}

        # ── 1. market_data 行情特征 ──
        cur.execute('SELECT * FROM market_data WHERE stock_code=%s', (code,))
        row = cur.fetchone()
        md_issue_date = None
        price_series = None
        if row:
            d = row
            md_issue_date = d.get('issue_date')
            ps_raw = d.get('price_series')
            if ps_raw:
                try:
                    price_series = json.loads(ps_raw)
                except Exception:
                    price_series = None

        if row and md_issue_date == issue_date:
            # ✅ 报价日匹配，直接用预计算的特征
            d = row
            for w in ['20d', '60d', '120d', '250d']:
                f[f'波动率_{w}'] = d.get(f'volatility_{w}')
                f[f'年化收益_{w}'] = d.get(f'annual_return_{w}')
                f[f'区间收益_{w}'] = d.get(f'period_return_{w}')
                f[f'胜率_{w}'] = d.get(f'win_rate_{w}')
            for ma in ['20', '30', '60', '120', '250']:
                f[f'MA{ma}'] = d.get(f'ma_{ma}')
            f['当前价'] = d.get('current_price')
            f['均价_all'] = d.get('avg_price_all')
            f['中位价'] = d.get('median_price')
            f['价格标准差'] = d.get('price_std')
            f['漂移率'] = d.get('drift')
            f['波动率'] = d.get('volatility')
            f['数据天数'] = d.get('total_days')
            f['报价日_md'] = d.get('issue_date')
            f['邀请日'] = d.get('invitation_date')
            f['最新交易日'] = d.get('latest_trading_date')
            f['行情_时间匹配'] = 1
        elif row and price_series and len(price_series) > 20:
            # ❌ 报价日不匹配（多次定增的早期事件），从 price_series 回算
            # 需要截取到报价日对应的位置
            # market_data 的 total_days 和 issue_date 告诉我们 series 有多长
            # series 最后一个 = md_issue_date 的价格
            # 我们需要估算 issue_date 对应的 series 索引
            # 方法: 用 issue_date_locked 获取报价日价格，在 series 中找到最接近的位置
            cur.execute(
                'SELECT issue_date_price FROM issue_date_locked WHERE stock_code=%s AND issue_date=%s',
                (code, issue_date)
            )
            il_row = cur.fetchone()
            target_price = il_row['issue_date_price'] if il_row else None

            cutoff_idx = len(price_series)  # 默认用全部
            if target_price is not None:
                prices_arr = np.array(price_series, dtype=float)
                diffs = np.abs(prices_arr - float(target_price))
                # 在后半段找最接近的
                search_start = max(0, len(prices_arr) // 2)
                cutoff_idx = search_start + np.argmin(diffs[search_start:])
                cutoff_idx = cutoff_idx + 1  # 包含该位置

            sub_series = price_series[:cutoff_idx]
            calc = _calc_features_from_series(sub_series)
            if calc:
                f.update(calc)
                f['行情_时间匹配'] = 0  # 标记为回算
                mismatch_count += 1
        else:
            f['行情_时间匹配'] = -1  # 无数据

        # ── 2. relative_valuation 估值特征 ──
        cur.execute('SELECT * FROM relative_valuation WHERE stock_code=%s', (code,))
        row = cur.fetchone()
        if row:
            d = row
            f['个股PE'] = d.get('current_pe')
            f['个股PB'] = d.get('current_pb')
            f['个股PS'] = d.get('current_ps')
            f['行业PE'] = d.get('sw_index_pe')
            f['行业PB'] = d.get('sw_index_pb')
            f['行业PS'] = d.get('sw_index_ps')
            f['行业代码'] = d.get('target_index_code', '')
            f['行业名称'] = d.get('target_industry_l3', '')

        # ── 3. historical_fcf FCF特征 (PIT year≤pit_year) ── 复用 _fcf_pit_cols(与回测单源)
        f.update(_fcf_pit_cols(code, issue_date, fcf_cache))

        # ── 4. issue_date_locked 锁定价 (按 报价日 精确匹配) ──
        cur.execute(
            'SELECT * FROM issue_date_locked WHERE stock_code=%s AND issue_date=%s',
            (code, issue_date)
        )
        row = cur.fetchone()
        if row:
            d = row
            f['报价日'] = d.get('issue_date')
            f['报价日价格'] = d.get('issue_date_price')
            f['报价日MA20'] = d.get('ma_20')
            f['锁定价当前价'] = d.get('current_price')

        # ── 5. placement_params 定增参数 (1230行) ──
        cur.execute('SELECT * FROM placement_params WHERE stock_code=%s', (code,))
        row = cur.fetchone()
        if row:
            d = row
            f['融资金额'] = d.get('financing_amount')
            f['锁定期'] = d.get('lockup_period')
            f['定价方式'] = d.get('pricing_method')
            f['溢价率'] = d.get('premium_rate')
            f['无风险利率'] = d.get('risk_free_rate')
            f['净资产负债表'] = d.get('net_assets')
            f['净债务'] = d.get('total_debt')
            f['净利润'] = d.get('net_income')
            f['营收增长率'] = d.get('revenue_growth')
            f['营业利润率'] = d.get('operating_margin')
            f['Beta'] = d.get('beta')

        # ── 6. screening_results 最新筛选结果 (4243行) ──
        # 注意: 溢价率/有效阈值数 已从Excel加载(更完整)，此处不再重复导出
        cur.execute(
            'SELECT * FROM screening_results WHERE stock_code=%s ORDER BY id DESC LIMIT 1', (code,)
        )
        row = cur.fetchone()
        if row:
            d = row
            # f['溢价率下限']/f['溢价率上限']/f['有效阈值数'] 由Excel提供(已标准化)
            f['step1通过'] = d.get('step1_pass')
            f['step2通过'] = d.get('step2_pass')
            f['step3通过'] = d.get('step3_pass')
            f['定增决策'] = d.get('decision')

        # ── 7. industry_data 行业行情数据 (1232行) ──
        cur.execute('SELECT * FROM industry_data WHERE stock_code=%s', (code,))
        row = cur.fetchone()
        if row:
            d = row
            f['行业级别'] = d.get('current_level')
            for w in ['20d', '60d', '120d', '250d']:
                f[f'行业波动率_{w}'] = d.get(f'volatility_{w}')
                f[f'行业年化收益_{w}'] = d.get(f'annual_return_{w}')
                f[f'行业区间收益_{w}'] = d.get(f'period_return_{w}')
                f[f'行业胜率_{w}'] = d.get(f'win_rate_{w}')
            f['行业MA20'] = d.get('ma_20')
            f['行业MA60'] = d.get('ma_60')
            f['行业MA120'] = d.get('ma_120')
            f['行业MA250'] = d.get('ma_250')
            f['行业漂移率'] = d.get('drift')
            f['行业波动率_总'] = d.get('volatility')
            f['行业总天数'] = d.get('total_days')
            # 行业分级信息
            f['sw_l1_code'] = d.get('sw_l1_code')
            f['sw_l1_name'] = d.get('sw_l1_name')
            f['sw_l2_code'] = d.get('sw_l2_code')
            f['sw_l2_name'] = d.get('sw_l2_name')
            f['sw_l3_code'] = d.get('sw_l3_code')
            f['sw_l3_name'] = d.get('sw_l3_name')

        # ── 8. peer_companies 同行公司数据 (17924行，聚合) ──
        cur.execute('SELECT * FROM peer_companies WHERE stock_code=%s', (code,))
        peers = cur.fetchall()
        if peers:
            peer_count = len(peers)
            f['同行公司数'] = peer_count
            peer_pe = [p['pe'] for p in peers if p.get('pe') is not None]
            peer_pb = [p['pb'] for p in peers if p.get('pb') is not None]
            peer_ps = [p['ps'] for p in peers if p.get('ps') is not None]
            peer_cap = [p['market_cap'] for p in peers if p.get('market_cap') is not None]
            if peer_pe:
                f['同行PE_均值'] = np.mean(peer_pe)
                f['同行PE_中位'] = np.median(peer_pe)
                f['同行PE_标准差'] = np.std(peer_pe)
            if peer_pb:
                f['同行PB_均值'] = np.mean(peer_pb)
                f['同行PB_中位'] = np.median(peer_pb)
            if peer_ps:
                f['同行PS_均值'] = np.mean(peer_ps)
                f['同行PS_中位'] = np.median(peer_ps)
            if peer_cap:
                f['同行市值_均值'] = np.mean(peer_cap)
                f['同行市值_中位'] = np.median(peer_cap)

        # ── 9. stocks 基础信息 (1250行) ──
        cur.execute('SELECT * FROM stocks WHERE stock_code=%s', (code,))
        row = cur.fetchone()
        if row:
            if 'sw_l1_name' not in f or not f.get('sw_l1_name'):
                f['sw_l1_code'] = row.get('sw_l1_code')
                f['sw_l1_name'] = row.get('sw_l1_name')
                f['sw_l2_code'] = row.get('sw_l2_code')
                f['sw_l2_name'] = row.get('sw_l2_name')
                f['sw_l3_code'] = row.get('sw_l3_code')
                f['sw_l3_name'] = row.get('sw_l3_name')

        features.append(f)

    conn.close()
    if mismatch_count > 0:
        print(f'   ⚠️ {mismatch_count}条样本的行情特征从price_series回算（报价日≠market_data日期）')
    return pd.DataFrame(features)


def load_scored_features(excel_path):
    """从scored Excel加载财务评分/子场景/涨跌幅"""
    df = pd.read_excel(excel_path, sheet_name='Sheet1')

    # 同一股票可能有多次定增（不同时间），均为独立样本，不去重
    dup_count = df['股票代码'].duplicated(keep=False).sum()
    if dup_count > 0:
        print(f'   ℹ️ 其中有{dup_count}条为同一股票多次定增（独立事件，保留）')

    result = pd.DataFrame()
    result['股票代码'] = df['股票代码']
    result['股票简称'] = df.get('股票简称', '')

    # 财务评分(按相对年份列)
    for prefix in ['总分', '评级', '盈利能力', '成长能力']:
        for c in df.columns:
            if c.startswith(f'{prefix}_') and c[len(prefix)+1:] in ['T-4','T-3','T-2','T-1','T']:
                result[c] = df[c]
            elif c.startswith(f'{prefix}_') and c[len(prefix)+1:].isdigit():
                result[c] = df[c]  # 兼容旧格式(绝对年份)

    # 趋势
    for c in ['总分_斜率', '总分_趋势', '盈利能力_斜率', '盈利能力_趋势',
              '成长能力_斜率', '成长能力_趋势', '综合趋势']:
        if c in df.columns:
            result[c] = df[c]

    # 子场景
    for c in ['市场指数', '行业PE', '个股PE', 'DCF估值', '修正PE估值',
              '参数构造', '蒙特卡洛', '反向推算']:
        if c in df.columns:
            result[f'{c}_通过'] = df[c].astype(str).str.contains('✓').astype(int)

    # 子场景通过总数
    sub_cols = [f'{c}_通过' for c in ['市场指数', '行业PE', '个股PE', 'DCF估值', '修正PE估值',
                                       '参数构造', '蒙特卡洛', '反向推算'] if f'{c}_通过' in result.columns]
    if sub_cols:
        result['子场景通过数'] = result[sub_cols].sum(axis=1)

    # 行业
    for c in ['一级行业', '二级行业', '三级行业']:
        if c in df.columns:
            result[c] = df[c]

    # 定增决策
    if '定增决策' in df.columns:
        result['定增建议参与'] = df['定增决策'].astype(str).str.contains('建议参与').astype(int)

    # 有效阈值数(从Excel，已标准化为0-1小数的溢价率，优先于DB版)
    if '有效阈值数' in df.columns:
        result['有效阈值数'] = df['有效阈值数']

    # 溢价率(从Excel，百分比字符串→小数，与screening_results的premium_min/max同义但更完整)
    for col in ['溢价率下限', '溢价率上限']:
        if col in df.columns:
            result[col] = pd.to_numeric(
                df[col].astype(str).str.replace('%', ''), errors='coerce'
            ) / 100

    # 标签: 7个月涨跌幅
    if '7个月后涨跌幅' in df.columns:
        ret = pd.to_numeric(
            df['7个月后涨跌幅'].astype(str).str.replace('%', '').str.replace('+', ''),
            errors='coerce'
        )
        result['7个月涨跌幅'] = ret
        result['标签_盈利_0'] = (ret > 0).astype(int)          # 盈利(>0)
        result['标签_盈利_-10'] = (ret > -10).astype(int)      # 盈利(>-10%)
        result['标签_盈利_-20'] = (ret > -20).astype(int)      # 盈利(>-20%)

    # 报价日价格(Excel版仅作fallback，DB版issue_date_locked.issue_date_price优先)
    # 不再单独导出，由load_db_features()提供

    # 7个月后价格
    if '7个月后价格' in df.columns:
        result['7个月后价格'] = pd.to_numeric(df['7个月后价格'], errors='coerce')

    # 最终结论
    if '最终结论' in df.columns:
        result['最终结论'] = df['最终结论']

    # 报价日(Excel版，用于构造sample_keys匹配DB，同时作为主列输出)
    if '报价日' in df.columns:
        result['报价日'] = df['报价日']

    return result


def load_scored_features_from_db(sample_keys):
    """从 investment_valuation DB 加载财务评分/子场景/标签（替代 Excel）

    Args:
        sample_keys: list of (stock_code, issue_date) 元组
    """
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')

    stock_codes = list(set(code for code, _ in sample_keys))
    codes_str = ','.join([f"'{c}'" for c in stock_codes])

    # 1. 从 placement_evaluation 获取评估记录
    pe_df = pd.read_sql(
        f'SELECT * FROM placement_evaluation WHERE stock_code IN ({codes_str})', conn
    )

    # 2. 从 company_annual_scores 获取年度评分
    cas_df = pd.read_sql(
        f'SELECT * FROM company_annual_scores WHERE stock_code IN ({codes_str}) ORDER BY report_year DESC', conn
    )
    conn.close()

    if pe_df.empty:
        print(f'  DB评分数据: 无匹配记录')
        return pd.DataFrame()

    # 构建结果: 每条 sample_key 一行
    results = []
    for code, issue_date in sample_keys:
        r = {'股票代码': code}

        # 匹配 placement_evaluation (按 stock_code + issue_date)
        pe_match = pe_df[pe_df['stock_code'] == code]
        if issue_date:
            pe_match_exact = pe_match[pe_match['issue_date'] == str(issue_date)]
            if not pe_match_exact.empty:
                pe_match = pe_match_exact
            elif not pe_match.empty:
                pe_match = pe_match.iloc[[0]]  # fallback: 取第一条
        else:
            pe_match = pe_match.iloc[[0]] if not pe_match.empty else pd.DataFrame()

        if not pe_match.empty:
            pe = pe_match.iloc[0]
            r['股票简称'] = pe.get('stock_name', '')
            r['报价日'] = pe.get('issue_date')
            r['报价日价格'] = pe.get('issue_date_price')

            # 子场景
            sub_map = {
                'sub_market_index': '市场指数', 'sub_industry_pe': '行业PE', 'sub_stock_pe': '个股PE',
                'sub_dcf': 'DCF估值', 'sub_adj_pe': '修正PE估值', 'sub_param_build': '参数构造',
                'sub_monte_carlo': '蒙特卡洛', 'sub_reverse_calc': '反向推算',
            }
            for db_col, cn_name in sub_map.items():
                v = pe.get(db_col, 0)
                r[f'{cn_name}_通过'] = int(v) if v and not (isinstance(v, float) and pd.isna(v)) else 0

            sub_cols = [f'{v}_通过' for v in sub_map.values()]
            r['子场景通过数'] = sum(r.get(c, 0) for c in sub_cols)

            # 行业
            r['一级行业'] = pe.get('industry_l1') or ''
            r['二级行业'] = pe.get('industry_l2') or ''
            r['三级行业'] = pe.get('industry_l3') or ''

            # 趋势
            r['总分_斜率'] = pe.get('total_slope')
            r['总分_趋势'] = pe.get('total_trend')
            r['盈利能力_斜率'] = pe.get('profit_slope')
            r['盈利能力_趋势'] = pe.get('profit_trend')
            r['成长能力_斜率'] = pe.get('growth_slope')
            r['成长能力_趋势'] = pe.get('growth_trend')
            r['综合趋势'] = pe.get('combined_trend')

            # 筛选
            r['有效阈值数'] = pe.get('valid_thresholds')
            r['溢价率下限'] = pe.get('premium_min')
            r['溢价率上限'] = pe.get('premium_max')
            r['定增决策'] = pe.get('decision')
            r['定增建议参与'] = 1 if '建议参与' in str(pe.get('decision', '')) else 0

            # 定增结构(东方财富 RPT_SEO_DETAIL + 易米主表; domain 核心特征源)
            #   raw 字段 → derive_features 算干净比率(折价率/稀释率/募集市值比/大股东参与/锁定期)
            #   绝对值(发行价/增发数/募资/股本)不可跨公司比较, 由 feature_exclusions 剔除, 仅作衍生原料
            r['定增_发行价'] = pe.get('em_issue_price')
            r['定增_增发数量'] = pe.get('em_issue_num')
            r['定增_募资总额'] = pe.get('em_raise_total')
            r['定增_发行前股本'] = pe.get('em_share_before')
            r['定增_发行后股本'] = pe.get('em_share_after')
            r['定增_发行对象'] = pe.get('em_issue_object')
            r['定增_定价原则'] = pe.get('em_price_principle')
            r['定增_发行方式'] = pe.get('em_issue_way')
            r['定增_解禁日'] = pe.get('pp_unlock_date')
            r['定增_承销商'] = pe.get('pp_underwriter')

            # 筹码分布衍生(cyq_chips 2018起; 强信号: chip_concentration 7m IV=0.22)
            #   由 fetch_chip_distribution.py 在报价日 PIT 算好回填, 此处直接发射
            r['chip_winner_rate'] = pe.get('chip_winner_rate')       # 获利盘比例
            r['chip_avg_cost_dev'] = pe.get('chip_avg_cost_dev')   # 平均成本/现价-1
            r['chip_concentration'] = pe.get('chip_concentration') # 筹码集中度 HHI
            r['chip_peak_dev'] = pe.get('chip_peak_dev')           # 筹码峰/现价-1
            r['chip_cost_spread'] = pe.get('chip_cost_spread')     # 成本离散 P75-P25/现价

            # 资金流 + 北向(moneyflow/hk_hold; fetch_factors capitalflow 回填) — 特征
            for _c in ('mf_main_net_ratio_5d', 'mf_main_net_ratio_20d', 'mf_net_mf_ratio_20d',
                       'mf_main_mom', 'mf_sm_net_ratio_20d', 'nb_hold_ratio',
                       'nb_hold_chg_20d', 'nb_hold_chg_60d'):
                r[_c] = pe.get(_c)
            # SMC 聪明钱(日/W/M; fetch_factors smc 回填; 周月线配月标签) — 特征
            # 通用发射: 自动发所有 smc_* 列(新增 smc_ote/smc_liqvoid 等免改此处)
            for _k, _v in pe.items():
                if isinstance(_k, str) and _k.startswith('smc_'):
                    r[_k] = _v
            # 业绩超预期 SUE(fetch_factors sue 回填; forecast/express/income PIT)
            # 通用发射: 自动发所有 sue_* 列(sue_yoy/zscore/beat/recency_d)
            for _k, _v in pe.items():
                if isinstance(_k, str) and _k.startswith('sue_'):
                    r[_k] = _v
            # 超额收益原始(compute_labels excess 回填; 目标变量→feature_exclusions 排除) + _0标签(跑赢基准)
            for _h in (1, 3, 7):
                for _b in ('mkt', 'ind'):
                    _e = pe.get(f'excess_{_b}_{_h}m')
                    r[f'excess_{_b}_{_h}m'] = _e
                    if pd.notna(_e):
                        r[f'标签_超{_b}_{_h}m_0'] = int(float(_e) > 0)
            # 短线收益原始(compute_labels shortterm 回填; 目标变量→排除) + _0标签
            for _w, _wn in (('1w', '1周'), ('2w', '2周'), ('4w', '4周')):
                _sr = pe.get(f'return_{_w}')
                r[f'{_wn}涨跌幅'] = _sr
                if pd.notna(_sr):
                    r[f'标签_短_{_w}_0'] = int(float(_sr) > 0)

            # 标签: 各期限涨跌幅 + 阈值标签 + 灰度剔除极性标签
            #   灰度区 [-20%, +10%] 剔除(NaN): 只留明显赢家(>+10%) vs 明显输家(<-20%)
            for _h in (1, 3, 6, 7, 12):
                ret = pe.get(f'return_{_h}m')
                if pd.notna(ret):
                    ret = float(ret)
                    r[f'{_h}个月涨跌幅'] = ret
                    r[f'标签_盈利_0_{_h}m'] = int(ret > 0)
                    r[f'标签_盈利_-10_{_h}m'] = int(ret > -10)
                    r[f'标签_盈利_-20_{_h}m'] = int(ret > -20)
                    # 灰度剔除极性: >+10=1(赢), <-20=0(输), 区间内=NaN(灰度, 训练时丢弃)
                    if ret > 10:
                        r[f'标签_极性_灰度剔除_{_h}m'] = 1
                    elif ret < -20:
                        r[f'标签_极性_灰度剔除_{_h}m'] = 0
                    else:
                        r[f'标签_极性_灰度剔除_{_h}m'] = np.nan
                r[f'{_h}个月后价格'] = pe.get(f'price_{_h}m')
            # 7m 向后兼容别名(默认标签口径, predict/registry 依赖)
            if pd.notna(pe.get('return_7m')):
                _r7 = float(pe['return_7m'])
                r['7个月涨跌幅'] = _r7
                r['标签_盈利_0'] = int(_r7 > 0)
                r['标签_盈利_-10'] = int(_r7 > -10)
                r['标签_盈利_-20'] = int(_r7 > -20)
            r['7个月后价格'] = pe.get('price_7m')
            r['最终结论'] = pe.get('final_conclusion')

        # 年度评分: 从 company_annual_scores 按报价日 PIT 回溯(最近可得期回退)。
        # 旧逻辑锚 base_year=报价日年 + report_year==base_year 取 T → 但当年年报要到次年
        # 4/30 才披露, T 槽几乎全空(实测 盈利能力_T 非空率仅 33.6%)。现按 PIT 法则:
        # 报价日年=Y, 5月起(Y-1)年报已披露可用, 否则只到(Y-2); T=≤该上限的最近可得年报年,
        # T-1..T-4 相对 T 顺延(保证 delta 的两年都已披露且存在)。不触碰未披露的当年/次年行→无泄漏。
        code_scores = cas_df[cas_df['stock_code'] == code].sort_values('report_year', ascending=False)

        # 报价日 → (基准年, 月份)
        issue_s = str(issue_date) if issue_date is not None else ''
        base_year = None
        if issue_s and issue_s not in ('', 'nan', 'None'):
            try:
                base_year = int(issue_s[:4])
            except (ValueError, TypeError):
                base_year = None
        if base_year is None and not code_scores.empty:
            base_year = int(code_scores.iloc[0]['report_year'])
        issue_month = int(issue_s[4:6]) if len(issue_s) >= 6 and issue_s[4:6].isdigit() else 12

        if base_year and not code_scores.empty:
            # PIT 上限: 法定年报披露截止 4/30 → 报价日 5月起上一年报可得, 否则再往前一年
            pit_max = base_year - 1 if issue_month >= 5 else base_year - 2
            avail = code_scores[code_scores['report_year'] <= pit_max]
            if not avail.empty:
                T_year = int(avail['report_year'].max())   # 最近可得(PIT)年报年
                labels = ['T-4', 'T-3', 'T-2', 'T-1', 'T']
                for i, label in enumerate(labels):
                    year = T_year - 4 + i
                    year_row = code_scores[code_scores['report_year'] == year]
                    if not year_row.empty:
                        yr = year_row.iloc[0]
                        for metric, col in [('总分', 'total_score'), ('评级', 'rating'),
                                            ('盈利能力', 'profitability'), ('成长能力', 'growth')]:
                            r[f'{metric}_{label}'] = yr.get(col)

        results.append(r)

    result_df = pd.DataFrame(results)
    # 超额/短线 gray 极性标签(p75/p25 分位, 全样本算) — compute_labels 只存原始, 标签在此统一发射
    for _h in (1, 3, 7):
        for _b in ('mkt', 'ind'):
            _col = f'excess_{_b}_{_h}m'
            if _col in result_df.columns:
                _s = pd.to_numeric(result_df[_col], errors='coerce')
                _p25, _p75 = _s.quantile(0.25), _s.quantile(0.75)
                result_df[f'标签_超{_b}_{_h}m_gray'] = np.where(_s > _p75, 1, np.where(_s < _p25, 0, np.nan))
    for _w, _wn in (('1w', '1周'), ('2w', '2周'), ('4w', '4周')):
        if f'{_wn}涨跌幅' in result_df.columns:
            _s = pd.to_numeric(result_df[f'{_wn}涨跌幅'], errors='coerce')
            _p25, _p75 = _s.quantile(0.25), _s.quantile(0.75)
            result_df[f'标签_短_{_w}_gray'] = np.where(_s > _p75, 1, np.where(_s < _p25, 0, np.nan))
    n = len(result_df)
    print(f'  DB评分数据: {n}条 (placement_evaluation + company_annual_scores)')
    return result_df


def load_financial_ratios(sample_keys):
    """从 investment_valuation.financial_indicators 按【报价日 point-in-time】加载 27 个财务比率。

    对每个 (code, issue_date) 取 ann_date <= 报价日 的最近一期年报(无未来财报泄漏)；
    ann_date 缺失时退化为 report_year <= 报价日年。返回行对齐 sample_keys 的 DataFrame
    (27 个中文比率列，不含 stock_code/财报年份，供与 scored/db_feats 行对齐 concat)。

    Args:
        sample_keys: list[(stock_code, issue_date)] —— 与 scored 行序一致
    """
    ratio_cols = {
        'current_ratio': '流动比率', 'quick_ratio': '速动比率',
        'inv_turn': '存货周转率', 'ar_turn': '应收账款周转率',
        'ca_turn': '流动资产周转率', 'assets_turn': '总资产周转率',
        'roa': 'ROA', 'npta': '总资产净利率',
        'roe': 'ROE', 'roe_dt': 'ROE摊薄',
        'cash_to_liqdebt': '现金流动负债比', 'cash_to_liqdebt_withinterest': '现金利息负债比',
        'netprofit_margin': '净利率', 'grossprofit_margin': '毛利率',
        'debt_to_assets': '资产负债率', 'int_to_talcap': '利息资本化率',
        'debt_to_eqt': '产权比率', 'ebit_to_interest': '已获利息倍数',
        'rd_exp_ratio': '研发费用率',
        'op_yoy': '营收增长', 'ebt_yoy': '利润增长',
        'netprofit_yoy': '净利增长', 'dt_netprofit_yoy': '扣非净利增长',
        'roe_yoy': 'ROE增长', 'tr_yoy': '总营收增长',
        'or_yoy': '营收增长2', 'equity_yoy': '净资产增长',
    }
    fields = list(ratio_cols.keys())
    cn_cols = list(ratio_cols.values())
    if not sample_keys:
        return pd.DataFrame(columns=cn_cols)

    codes = list({str(c) for c, _ in sample_keys})
    try:
        conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                               database='investment_valuation', charset='utf8mb4')
        codes_str = ','.join([f"'{c}'" for c in codes])
        select_cols = ', '.join(['stock_code', 'report_year', 'ann_date'] + fields)
        df_all = pd.read_sql(
            f'SELECT {select_cols} FROM financial_indicators WHERE stock_code IN ({codes_str})',
            conn)
        conn.close()
    except Exception as e:
        print(f'  财务比率 PIT: 不可用({e})')
        return pd.DataFrame(columns=cn_cols)

    # 预按 stock_code 分组 + 每股排序/数组化一次(免逐键 sort_values/astype/Series索引; 53万样本必需)
    grouped_prep = {}
    for c, g in df_all.groupby('stock_code'):
        gs = g.sort_values('ann_date', ascending=False, na_position='last')
        grouped_prep[str(c)] = (gs['ann_date'].astype(str).values, gs['report_year'].values,
                                {f: gs[f].values for f in fields})
    rows, n_hit = [], 0
    for code, issue_date in sample_keys:
        code = str(code)
        ids = str(issue_date) if issue_date is not None else ''
        if len(ids) >= 8 and ids[:8].isdigit():
            id8, yr = ids[:8], int(ids[:4])
        else:
            id8, yr = '99999999', 9999  # 报价日缺失 → 取最新(无 PIT 约束)
        prep = grouped_prep.get(code)
        if prep is None:
            rows.append({}); continue
        ad_arr, ry_arr, vals = prep
        # ann_date 非空且 ≤报价日: ad_arr 排序 desc('nan'排末尾且 'nan'>id8 自动排除) → 首个True=最新有效
        m1 = ad_arr <= id8
        if m1.any():
            idx = int(np.argmax(m1))
        else:
            m2 = (ad_arr == 'nan') & (ry_arr <= yr)   # 退化: ann_date空 且 report_year≤报价日年
            if not m2.any():
                rows.append({}); continue
            idx = int(np.argmax(m2))
        n_hit += 1
        rows.append({ratio_cols[f]: vals[f][idx] for f in fields if pd.notna(vals[f][idx])})

    out = pd.DataFrame(rows)
    for cn in cn_cols:
        if cn not in out.columns:
            out[cn] = np.nan
    print(f'  财务比率 PIT(ann_date≤报价日): 命中 {n_hit}/{len(sample_keys)} 样本')
    return out[cn_cols]


def load_financial_statements(stock_codes):
    """从 fund_risk_control MySQL 加载资产负债表/利润表/现金流量表的关键指标"""
    try:
        conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                               database='fund_risk_control', charset='utf8mb4')
        codes_str = ','.join([f"'{c}'" for c in stock_codes])

        # ── 资产负债表关键指标 ──
        bs_df = pd.read_sql(
            f"""SELECT ts_code, report_year,
                total_assets, total_cur_assets, total_nca, total_liab, total_cur_liab, total_ncl,
                total_hldr_eqy_exc_min_int, money_cap, inventories, accounts_receiv,
                goodwill, intan_assets, fix_assets, total_share,
                st_borr, lt_borr, bond_payable
                FROM balance_sheet WHERE ts_code IN ({codes_str})""", conn
        )

        # ── 利润表关键指标 ──
        is_df = pd.read_sql(
            f"""SELECT ts_code, report_year,
                total_revenue, revenue, operate_profit, total_profit, n_income,
                n_income_attr_p, sell_exp, admin_exp, fin_exp, rd_exp,
                basic_eps, ebit, ebitda
                FROM income_statement WHERE ts_code IN ({codes_str})""", conn
        )

        # ── 现金流量表关键指标 ──
        cf_df = pd.read_sql(
            f"""SELECT ts_code, report_year,
                n_cashflow_act, n_cash_inv_act, subtotal_operate_cash_flow,
                invest_cash_flow, finance_cash_flow, net_profit
                FROM cash_flow WHERE ts_code IN ({codes_str})""", conn
        )

        conn.close()

        result_dfs = []

        # 处理资产负债表
        if not bs_df.empty:
            bs_df = bs_df.sort_values('report_year', ascending=False).groupby('ts_code').first().reset_index()
            # 计算派生指标
            if 'total_assets' in bs_df.columns and 'total_liab' in bs_df.columns:
                bs_df['资产负债比'] = pd.to_numeric(bs_df['total_assets'], errors='coerce') / \
                                     pd.to_numeric(bs_df['total_liab'], errors='coerce').replace(0, np.nan)
            if 'total_cur_assets' in bs_df.columns and 'total_cur_liab' in bs_df.columns:
                bs_df['流动比率_bs'] = pd.to_numeric(bs_df['total_cur_assets'], errors='coerce') / \
                                       pd.to_numeric(bs_df['total_cur_liab'], errors='coerce').replace(0, np.nan)
            if 'goodwill' in bs_df.columns and 'total_assets' in bs_df.columns:
                bs_df['商誉占总资产比'] = pd.to_numeric(bs_df['goodwill'], errors='coerce') / \
                                         pd.to_numeric(bs_df['total_assets'], errors='coerce').replace(0, np.nan)
            if 'intan_assets' in bs_df.columns and 'total_assets' in bs_df.columns:
                bs_df['无形资产占比'] = pd.to_numeric(bs_df['intan_assets'], errors='coerce') / \
                                        pd.to_numeric(bs_df['total_assets'], errors='coerce').replace(0, np.nan)
            # 重命名
            rename_bs = {
                'total_assets': '总资产', 'total_cur_assets': '流动资产', 'total_nca': '非流动资产',
                'total_liab': '总负债', 'total_cur_liab': '流动负债', 'total_ncl': '非流动负债',
                'total_hldr_eqy_exc_min_int': '股东权益', 'money_cap': '货币资金',
                'inventories': '存货', 'accounts_receiv': '应收账款',
                'goodwill': '商誉', 'intan_assets': '无形资产', 'fix_assets': '固定资产',
                'total_share': '总股本', 'st_borr': '短期借款', 'lt_borr': '长期借款',
                'bond_payable': '应付债券',
            }
            rename_map = {k: v for k, v in rename_bs.items() if k in bs_df.columns}
            bs_df = bs_df.rename(columns=rename_map)
            bs_df = bs_df.rename(columns={'ts_code': '股票代码'})
            result_dfs.append(bs_df)
            print(f'  资产负债表: 匹配到 {len(bs_df)} 家公司')

        # 处理利润表
        if not is_df.empty:
            is_df = is_df.sort_values('report_year', ascending=False).groupby('ts_code').first().reset_index()
            # 派生指标
            if 'operate_profit' in is_df.columns and 'total_revenue' in is_df.columns:
                is_df['营业利润率_is'] = pd.to_numeric(is_df['operate_profit'], errors='coerce') / \
                                          pd.to_numeric(is_df['total_revenue'], errors='coerce').replace(0, np.nan)
            if 'n_income' in is_df.columns and 'total_revenue' in is_df.columns:
                is_df['净利润率_is'] = pd.to_numeric(is_df['n_income'], errors='coerce') / \
                                       pd.to_numeric(is_df['total_revenue'], errors='coerce').replace(0, np.nan)
            if 'sell_exp' in is_df.columns and 'total_revenue' in is_df.columns:
                is_df['销售费用率'] = pd.to_numeric(is_df['sell_exp'], errors='coerce') / \
                                      pd.to_numeric(is_df['total_revenue'], errors='coerce').replace(0, np.nan)
            if 'admin_exp' in is_df.columns and 'total_revenue' in is_df.columns:
                is_df['管理费用率'] = pd.to_numeric(is_df['admin_exp'], errors='coerce') / \
                                      pd.to_numeric(is_df['total_revenue'], errors='coerce').replace(0, np.nan)
            if 'rd_exp' in is_df.columns and 'total_revenue' in is_df.columns:
                is_df['研发费用率_is'] = pd.to_numeric(is_df['rd_exp'], errors='coerce') / \
                                          pd.to_numeric(is_df['total_revenue'], errors='coerce').replace(0, np.nan)
            rename_is = {
                'total_revenue': '营业总收入', 'revenue': '营业收入', 'operate_profit': '营业利润',
                'total_profit': '利润总额', 'n_income': '净利润', 'n_income_attr_p': '归母净利润',
                'basic_eps': '基本EPS', 'ebit': 'EBIT', 'ebitda': 'EBITDA',
            }
            rename_map = {k: v for k, v in rename_is.items() if k in is_df.columns}
            is_df = is_df.rename(columns=rename_map)
            is_df = is_df.rename(columns={'ts_code': '股票代码'})
            result_dfs.append(is_df)
            print(f'  利润表: 匹配到 {len(is_df)} 家公司')

        # 处理现金流量表
        if not cf_df.empty:
            cf_df = cf_df.sort_values('report_year', ascending=False).groupby('ts_code').first().reset_index()
            rename_cf = {
                'n_cashflow_act': '经营现金流', 'n_cash_inv_act': '投资现金流',
                'subtotal_operate_cash_flow': '经营活动现金流小计',
                'invest_cash_flow': '投资活动现金流', 'finance_cash_flow': '筹资活动现金流',
            }
            rename_map = {k: v for k, v in rename_cf.items() if k in cf_df.columns}
            cf_df = cf_df.rename(columns=rename_map)
            cf_df = cf_df.rename(columns={'ts_code': '股票代码'})
            result_dfs.append(cf_df)
            print(f'  现金流量表: 匹配到 {len(cf_df)} 家公司')

        if not result_dfs:
            return pd.DataFrame()

        # 合并三表
        merged = result_dfs[0]
        for dff in result_dfs[1:]:
            # 避免report_year列重复
            dup_cols = [c for c in dff.columns if c in merged.columns and c != '股票代码']
            if dup_cols:
                dff = dff.drop(columns=dup_cols)
            merged = merged.merge(dff, on='股票代码', how='outer')

        return merged

    except Exception as e:
        print(f'  财务报表: 不可用({e})')
        return pd.DataFrame()


def load_sample_keys_from_db():
    """从 placement_evaluation 取全部 (stock_code, issue_date) 作为样本键（整条流水线的脊柱）。

    Returns:
        list of (stock_code, issue_date_str) 元组，issue_date 形如 '20200109'
    """
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    df = pd.read_sql(
        "SELECT DISTINCT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND issue_date <> ''",
        conn,
    )
    conn.close()
    return [(r['stock_code'], str(r['issue_date'])) for _, r in df.iterrows()]


def main():
    parser = argparse.ArgumentParser(description='定增特征导出(全量)')
    parser.add_argument('excel_path', nargs='?', default=None,
                        help='(已废弃，评分/标签现从 DB 读取) 保留仅作兼容，不再使用')
    parser.add_argument('--output', default=None, help='输出文件路径(默认ml_training/data/features.parquet)')
    parser.add_argument('--no-mysql', action='store_true', help='跳过fund_risk_control数据')
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = args.output or os.path.join(script_dir, 'data', 'features.parquet')
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # ── 1. 从 placement_evaluation 取样本键 + 财务评分/标签 ──
    print('1. 从 DB 加载样本键 (placement_evaluation)...')
    sample_keys = load_sample_keys_from_db()
    print(f'   {len(sample_keys)} 条 (stock_code, issue_date)')
    if not sample_keys:
        print('   ⚠️ placement_evaluation 无数据，请先运行 scripts/data_pipeline/backfill_evaluations.py')
        sys.exit(1)

    print('   加载评分/标签 (placement_evaluation + company_annual_scores)...')
    scored = load_scored_features_from_db(sample_keys)
    print(f'   {len(scored)} 只股票')

    # ── 解禁日过滤：剔除解禁日(报价日+锁定期)晚于截止日的未解禁样本 ──
    # 这些定增尚未解禁，盈亏未落定，不能进训练集。DB 无解禁日字段，按标准
    # 6 个月锁定期推算；截止日按数据观测时点，需更新时改 UNLOCK_CUTOFF 即可。
    LOCKUP_MONTHS = 6
    # 截止日 = 取数当天(动态)：只保留解禁日 ≤ 今日 的样本(盈亏已落定)
    UNLOCK_CUTOFF = pd.Timestamp.now().normalize()
    _qd = pd.to_datetime(
        pd.to_numeric(scored['报价日'], errors='coerce').dropna().astype(int).astype(str),
        format='%Y%m%d', errors='coerce'
    ).reindex(scored.index)
    _unlock = _qd + pd.DateOffset(months=LOCKUP_MONTHS)
    _not_unlocked = (_unlock > UNLOCK_CUTOFF).fillna(False)
    if _not_unlocked.any():
        n_drop = int(_not_unlocked.sum())
        scored = scored[~_not_unlocked].reset_index(drop=True)
        print(f'   ⚠️ 解禁日过滤: 剔除 {n_drop} 个未解禁样本 '
              f'(报价日+{LOCKUP_MONTHS}月 > {UNLOCK_CUTOFF.date()})')
        # 同步重建 sample_keys，保持与 scored 行对齐(load_db_features 按行索引取)
        scored['报价日'] = pd.to_numeric(scored['报价日'], errors='coerce')  # 'nan'字符串→NaN, 防 int() 崩
        sample_keys = [
            (r['股票代码'], str(int(r['报价日'])) if pd.notna(r['报价日']) else '')
            for _, r in scored.iterrows()
        ]
    print(f'   过滤后 {len(scored)} 只股票')

    stock_codes = scored['股票代码'].tolist()

    # ── 2. 从investment_valuation加载全部DB特征 ──
    print('2. 加载investment_valuation DB特征...')
    # sample_keys 与 scored/load_db_features 三者行顺序一致，按行索引对齐
    db_feats = load_db_features(sample_keys)
    matched = db_feats['当前价'].notna().sum()
    print(f'   行情数据: {matched}/{len(stock_codes)}')
    time_match = db_feats.get('行情_时间匹配')
    if time_match is not None:
        exact = (time_match == 1).sum()
        recalc = (time_match == 0).sum()
        missing = (time_match == -1).sum()
        print(f'   行情时间匹配: 精确={exact}, 回算={recalc}, 无数据={missing}')
    matched_ind = db_feats['行业波动率_20d'].notna().sum()
    print(f'   行业数据: {matched_ind}/{len(stock_codes)}')
    matched_peer = db_feats['同行公司数'].notna().sum()
    print(f'   同行数据: {matched_peer}/{len(stock_codes)}')

    # ── 3. 从fund_risk_control加载财务比率（三表已弃用：fund_risk_control 仅7家数据，
    #        覆盖率0.43%，且与 financial_indicators API 比率表重复，全部 importance=0）──
    if not args.no_mysql:
        print('3. 加载财务比率(financial_indicators)...')
        ratio_feats = load_financial_ratios(sample_keys)
    else:
        ratio_feats = pd.DataFrame()

    # ── 4. 合并所有特征 ──
    print('4. 合并全部特征...')
    # 用行索引对齐（避免重复股票代码的笛卡尔积）
    scored = scored.reset_index(drop=True)
    db_feats = db_feats.reset_index(drop=True)
    # 去掉 db_feats 中与 scored 同名的列（如 股票代码/报价日/报价日价格/定增决策…），
    # 一律取 scored（来自 placement_evaluation，权威源）的版本，避免 concat 后出现重名列。
    _dup_cols = set(scored.columns) & set(db_feats.columns)
    if _dup_cols:
        print(f'   去重列(取scored版): {sorted(_dup_cols)}')
    db_feat_cols = [c for c in db_feats.columns if c not in _dup_cols]
    merged = pd.concat([scored, db_feats[db_feat_cols]], axis=1)

    # 财务比率已按 sample_keys 行对齐(PIT)，直接 concat
    if not ratio_feats.empty:
        merged = pd.concat([merged.reset_index(drop=True), ratio_feats.reset_index(drop=True)], axis=1)

    # ── 剔除死字段（importance=0 且 IV<0.02，模型从未使用）──
    # 来源：三表已不加载；以下是其余无预测力的标识/稀疏/重复字段
    EXCLUDE_FIELDS = {
        # 行业代码/时间标识（文本，本不该入模）
        'sw_l1_code', 'sw_l1_name', 'sw_l2_code', 'sw_l2_name', 'sw_l3_code', 'sw_l3_name',
        'report_year', '财报年份',
        # 子场景/筛选通过标记（近常数）
        '市场指数_通过', '参数构造_通过', '反向推算_通过', 'step1通过', 'step3通过',
        '定增建议参与', '有效阈值数', '行情_时间匹配',
        # 定增参数（IV=0）
        '溢价率', '溢价率下限', '锁定期', '融资金额', '无风险利率', 'Beta',
        # 稀疏 FCF 营运资金变动（全0/极稀疏）
        '营运资金变动_T', '营运资金变动_T1', '营运资金变动_T2', '营运资金变动_T3', '营运资金变动_T4',
        # 弱/重复比率
        '营业利润率', '营收增长率', '行业PS', '市场_above_MA250',
        '净债务', '净资产负债表',
    }
    drop_present = [c for c in EXCLUDE_FIELDS if c in merged.columns]
    if drop_present:
        merged = merged.drop(columns=drop_present)
        print(f'   剔除死字段: {len(drop_present)} 个')

    # 列数统计
    total_cols = len(merged.columns)
    label_cols = [c for c in merged.columns if c.startswith('标签_')]
    target_cols = ['7个月涨跌幅', '7个月后价格', '最终结论']
    feature_cols = [c for c in merged.columns if c not in label_cols + target_cols]
    print(f'   总特征数: {len(feature_cols)}, 标签数: {len(label_cols)}, 总列数: {total_cols}')

    # 数值型特征非空率统计
    num_cols = merged[feature_cols].select_dtypes(include=[np.number]).columns.tolist()
    print(f'   数值特征: {len(num_cols)}个')
    nonnull_rates = merged[num_cols].notna().mean()
    high_coverage = (nonnull_rates > 0.5).sum()
    print(f'   覆盖率>50%: {high_coverage}个')

    # ── 5. 保存 ──
    # 清理混合类型列：对疑似数值列确保不含字符串
    # 保留关键字符串列不动
    str_cols_keep = {'股票代码', '股票简称', '最终结论', '一级行业', '二级行业', '三级行业',
                     '定价方式', '定增决策', '行业代码', '行业名称'}

    def _is_str_col(name):
        """判定列是否应保留为字符串（文本判定/评级，不转数值）"""
        if name in str_cols_keep:
            return True
        # 趋势判定: 总分_趋势/盈利能力_趋势/成长能力_趋势/综合趋势 → 通过/不通过
        if name.endswith('_趋势') or name == '综合趋势':
            return True
        # 评级: 评级_T/评级_T-1/... → A/BBB/CCC
        if name.startswith('评级_'):
            return True
        return False

    for c in merged.columns:
        if _is_str_col(c):
            merged[c] = merged[c].astype(str)
        elif merged[c].dtype == object:
            # 尝试转为数值，失败的变NaN
            merged[c] = pd.to_numeric(merged[c], errors='coerce')

    merged.to_parquet(output_path, index=False)
    print(f'\n✅ 特征已保存: {output_path}')
    print(f'   {len(merged)} 行 × {total_cols} 列')

    csv_path = output_path.replace('.parquet', '.csv')
    merged.to_csv(csv_path, index=False)
    print(f'   CSV(查看用): {csv_path}')

    # 字段说明
    schema_path = os.path.join(os.path.dirname(output_path), 'features_schema.md')
    with open(schema_path, 'w', encoding='utf-8') as f:
        f.write('# 特征字段说明\n\n')
        f.write(f'总行数: {len(merged)}, 总列数: {total_cols}\n\n')

        groups = {
            '基础信息': [c for c in feature_cols if any(k in c for k in
                        ['股票代码', '股票简称', '报价日', '行业', '一级行业', '二级行业', '三级行业',
                         'sw_l1', 'sw_l2', 'sw_l3', '最新交易日', '邀请日'])],
            '行情特征': [c for c in num_cols if any(k in c for k in
                        ['波动率', '年化收益', '区间收益', '胜率', 'MA', '漂移', '换手',
                         '当前价', '数据天数', '均价', '中位价', '价格标准差']) and not c.startswith('行业')],
            '行业行情': [c for c in num_cols if c.startswith('行业')],
            '估值特征': [c for c in num_cols if any(k in c for k in
                        ['个股PE', '个股PB', '个股PS', '行业PE', '行业PB', '行业PS'])],
            '同行对比': [c for c in num_cols if c.startswith('同行')],
            'FCF特征': [c for c in num_cols if any(k in c for k in
                        ['营收', 'NOPAT', 'FCF', '营业利润_fcf', '净利润_fcf', '折旧', '资本支出',
                         '营运资金', 'FCF年份'])],
            '定增参数': [c for c in num_cols if any(k in c for k in
                        ['融资', '锁定', '净债', '净利润', '净资产负债', '营收增长率', '营业利润率',
                         'Beta', '溢价率', '无风险', '定价方式', '发行价'])],
            '筛选决策': [c for c in num_cols if any(k in c for k in
                        ['有效阈值', 'step', '定增决策', '定增建议'])],
            '财务评分': [c for c in feature_cols if any(k in c for k in
                        ['总分', '评级', '盈利能力', '成长能力', '斜率', '趋势', '综合', '子场景'])],
            '子场景': [c for c in feature_cols if '_通过' in c],
            '财务比率': [c for c in num_cols if c in
                       ['流动比率', '速动比率', '存货周转率', '应收账款周转率',
                        '总资产周转率', 'ROA', 'ROE', 'ROE摊薄',
                        '净利率', '毛利率', '资产负债率', '产权比率',
                        '已获利息倍数', '研发费用率', '营收增长', '净利增长',
                        '扣非净利增长', 'ROE增长', '净资产增长', '财报年份']],
            '资产负债表': [c for c in num_cols if c in
                       ['总资产', '流动资产', '非流动资产', '总负债', '流动负债', '非流动负债',
                        '股东权益', '货币资金', '存货_bs', '应收账款', '商誉', '无形资产',
                        '固定资产', '总股本', '短期借款', '长期借款', '应付债券',
                        '资产负债比', '流动比率_bs', '商誉占总资产比', '无形资产占比']],
            '利润表': [c for c in num_cols if c in
                     ['营业总收入', '营业收入', '营业利润', '利润总额', '净利润_is',
                      '归母净利润', '基本EPS', 'EBIT', 'EBITDA',
                      '营业利润率_is', '净利润率_is', '销售费用率', '管理费用率', '研发费用率_is']],
            '现金流量表': [c for c in num_cols if c in
                       ['经营现金流', '投资现金流', '经营活动现金流小计',
                        '投资活动现金流', '筹资活动现金流']],
            '标签': label_cols + target_cols,
        }

        for group_name, cols in groups.items():
            if cols:
                f.write(f'\n## {group_name} ({len(cols)}个)\n\n')
                f.write('| 字段 | 非空率 |\n|------|--------|\n')
                for c in cols:
                    rate = merged[c].notna().mean() * 100 if c in merged.columns else 0
                    f.write(f'| {c} | {rate:.1f}% |\n')
    print(f'   字段说明: {schema_path}')

    # ── 冻结快照入 DB 版本库(data 文件已 gitignore, 靠此还原) ──
    try:
        import db_dataset_store
        base_version = db_dataset_store.save_snapshot(
            merged, kind='base', label_config='7m',
            note=f'base features {len(merged)}x{total_cols}')
        print(f'   快照入DB: {base_version}')
        with open(schema_path, 'a', encoding='utf-8') as f:
            f.write(f'\n\n## 数据集快照\n\n基线快照版本: `{base_version}`\n'
                    f'还原: `python manage_snapshots.py restore {base_version} '
                    f'--out data/features.parquet`\n')
    except Exception as e:
        print(f'   ⚠️ 快照入库跳过(DB不可用): {e}')


if __name__ == '__main__':
    main()

# ═══════════════════════════════════════════════════════════════════════
# PIT 特征 loader(原 scripts/feature_loaders.py, 2026-06-26 并入: 唯一特征引擎, 消除双线分叉)
# 5 特殊特征(FCF_加速/总分_delta_2y/nb_hold_ratio/PB_vs_同行中位/sue_beat)PIT loader,
# 接受(codes,date)/(code,date)keys, PIT ≤date, 不 join placement_evaluation。定增/回测/未来共用。
# ═══════════════════════════════════════════════════════════════════════
_PKG_FL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PKG_FL not in sys.path: sys.path.insert(0, _PKG_FL)
if os.path.join(_PKG_FL, 'scripts') not in sys.path: sys.path.insert(0, os.path.join(_PKG_FL, 'scripts'))

_DB = dict(host='127.0.0.1', port=3306, user='root', password='',
           database='investment_valuation', charset='utf8mb4')


def _pit_max(date_yyyymmdd):
    """年报 PIT 年份: 月≥5 用上年(已披露), 否则前年。"""
    s = str(date_yyyymmdd)
    by, mo = int(s[:4]), int(s[4:6])
    return by - 1 if mo >= 5 else by - 2


def _conn():
    return pymysql.connect(**_DB)


def _pro():
    """共享 tushare pro_api(token 走 resolve_tushare_token, 不硬编码)。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    return ts.pro_api()


# ─────────────── 全历史预取缓存(FCF/总分/SUE, 避免逐股逐日查) ───────────────
_FCF_CACHE = {}     # code → DataFrame[year, fcf] 全历史(升序)
_SCORE_CACHE = {}   # code → DataFrame[report_year, total_score] 全历史(升序)
_SUE_CACHE = {}     # code → 披露时间线 DataFrame(forecast/express/income 合并, ann_date 升序) 或 None
# SUE 时间线落盘缓存: 历史披露是 PIT 固定的(永不改变), 落盘后回测重建 panel 免重取(500股≈15min→秒级)。
# 增量(未来新披露)用 refresh=True 重取覆盖; 否则只补磁盘+内存都没有的股。
_SUE_PARQ = os.path.join(_PKG_FL, 'ml_training', 'data', 'sue_timelines.parquet')


def _chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def prefetch_fcf_scores(codes):
    """一次性批量预取全 universe 的 FCF + 总分全历史(各 N 条 SQL, IN 分块 500),
    缓存到 _FCF_CACHE/_SCORE_CACHE。后续 load_fcf_accel/load_total_score_delta2y 只做 PIT 内存切片。
    把"500股×192月×2查询≈19万次"降到"500股×2≈1000次"(预取) + 内存切片(免费)。"""
    codes = [str(c) for c in codes if c is not None]
    conn = _conn()
    for c in [c for c in codes if c not in _FCF_CACHE]:
        pass   # 触发下面批量
    new_fcf = [c for c in codes if c not in _FCF_CACHE]
    new_sc = [c for c in codes if c not in _SCORE_CACHE]
    if new_fcf:
        for chunk in _chunks(new_fcf, 500):
            ph = ','.join([f"'{c}'" for c in chunk])
            df = pd.read_sql(f"SELECT stock_code, year, fcf FROM historical_fcf "
                             f"WHERE stock_code IN ({ph}) AND fcf IS NOT NULL", conn)
            for c, g in df.groupby('stock_code'):
                _FCF_CACHE[c] = g.sort_values('year').reset_index(drop=True)
    if new_sc:
        for chunk in _chunks(new_sc, 500):
            ph = ','.join([f"'{c}'" for c in chunk])
            df = pd.read_sql(f"SELECT stock_code, report_year, total_score FROM company_annual_scores "
                             f"WHERE stock_code IN ({ph}) AND total_score IS NOT NULL", conn)
            for c, g in df.groupby('stock_code'):
                _SCORE_CACHE[c] = g.sort_values('report_year').reset_index(drop=True)
    conn.close()
    print(f'    预取 FCF({len(_FCF_CACHE)}股) + 总分({len(_SCORE_CACHE)}股) 全历史缓存')


def _load_sue_disk():
    """从 _SUE_PARQ 载入已预取的 SUE 时间线 → {code: DataFrame 或 None}。
    空时间线存为 _has_data=0 占位行 → 载回 None(避免每轮重取空股的 3 次 API)。
    无文件返回 {}。"""
    if not os.path.exists(_SUE_PARQ):
        return {}
    try:
        df = pd.read_parquet(_SUE_PARQ)
    except Exception:
        return {}
    if df is None or df.empty or 'stock_code' not in df.columns:
        return {}
    out = {}
    for c, g in df.groupby('stock_code'):
        c = str(c)
        if int(g['_has_data'].iloc[0]) == 0:
            out[c] = None
        else:
            out[c] = (g[g['_has_data'] == 1]
                      .drop(columns=['stock_code', '_has_data'], errors='ignore')
                      .reset_index(drop=True))
    return out


def _save_sue_disk():
    """把 _SUE_CACHE 落盘 parquet(增量合并: 保留磁盘已有股, 本次新取的同股覆盖)。
    空时间线写 1 行 _has_data=0 占位; 有效时间线每披露 1 行 _has_data=1。"""
    frames = []
    for c, g in _SUE_CACHE.items():
        c = str(c)
        if g is None or (hasattr(g, 'empty') and g.empty):
            frames.append(pd.DataFrame({'stock_code': [c], '_has_data': [0]}))
            continue
        gg = g.copy()
        gg['stock_code'] = c
        gg['_has_data'] = 1
        frames.append(gg)
    if not frames:
        return
    new_df = pd.concat(frames, ignore_index=True)
    new_codes = set(new_df['stock_code'])
    if os.path.exists(_SUE_PARQ):
        try:
            old = pd.read_parquet(_SUE_PARQ)
            old = old[~old['stock_code'].astype(str).isin(new_codes)]
            new_df = pd.concat([old, new_df], ignore_index=True)
        except Exception:
            pass
    new_df.to_parquet(_SUE_PARQ, index=False)


def prefetch_sue_timelines(codes, refresh=False):
    """预取 forecast/express/income 披露时间线(每股 3 次 tushare API), 缓存供 load_sue_beat PIT 切片。
    复用 fetch_factors._build_disclosure_timeline; 无 DB 依赖(纯 API), 任意股可取。

    历史披露是 PIT 固定的 → 落盘 _SUE_PARQ 后, 回测重建 panel 不再重取(省 ~3 API/股)。
    refresh=True: 忽略磁盘, 全量重取并覆盖(增量新披露用)。"""
    from data_pipeline.fetch_factors import _build_disclosure_timeline
    # 1) 先载入磁盘缓存(除非 refresh)
    if refresh:
        _SUE_CACHE.clear()
    elif not _SUE_CACHE:
        _SUE_CACHE.update(_load_sue_disk())
        if _SUE_CACHE:
            ok0 = sum(1 for v in _SUE_CACHE.values() if v is not None)
            print(f'    载入磁盘 SUE 时间线: {len(_SUE_CACHE)} 股({ok0} 有数据)')
    # 2) 只对磁盘+内存都没有的逐股取 API
    new = [str(c) for c in codes if str(c) not in _SUE_CACHE]
    if new:
        pro = _pro()
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import time
        print(f'    并发取 SUE 时间线 {len(new)} 股(每股 3 API, max_workers=12, 限流重试)...')
        def _work(c):
            for attempt in range(3):
                try:
                    return c, _build_disclosure_timeline(pro, c)
                except Exception:
                    time.sleep(1.0 * (attempt + 1))   # 限流/网络 → 退避重试
            return c, None
        fetched = 0
        with ThreadPoolExecutor(max_workers=12) as ex:
            futs = [ex.submit(_work, c) for c in new]
            for n, fut in enumerate(as_completed(futs), 1):
                c, tl = fut.result()
                _SUE_CACHE[c] = tl            # GIL 下不同 key 并发写安全
                if tl is not None:
                    fetched += 1
                if n % 100 == 0:
                    print(f'    SUE {n}/{len(new)} (有数据 {fetched})')
        if fetched:
            _save_sue_disk()
            print(f'    落盘 {fetched} 条新 SUE 时间线 → {os.path.basename(_SUE_PARQ)}')
    ok = sum(1 for v in _SUE_CACHE.values() if v is not None)
    print(f'    预取 SUE 披露时间线: {ok}/{len(_SUE_CACHE)} 股有数据'
          + ('(本次重取)' if refresh else f'(磁盘缓存 {sum(1 for v in _load_sue_disk().values() if v is not None)} 有效)'))


# ─────────────── 1. FCF_加速 (historical_fcf PIT 内存切片) ───────────────
def load_fcf_accel(keys):
    """keys=[(code,date)]。用 _FCF_CACHE 全历史, PIT 切 year≤pit_max 取 T/T1/T2 → YoY 加速度。
    需先 prefetch_fcf_scores 预热(回测 run() 开头调一次)。"""
    if not keys:
        return pd.DataFrame(columns=['FCF_加速'])
    out = {}
    for code, date in keys:
        g = _FCF_CACHE.get(code)
        if g is None or g.empty:
            out[code] = np.nan; continue
        sub = g[g['year'] <= _pit_max(date)]
        rs = sub['fcf'].tail(3).values      # 升序, 末3个 = T/T1/T2(最大年)
        if len(rs) >= 2:
            yoy_t = (rs[-1] - rs[-2]) / abs(rs[-2]) if rs[-2] else np.nan
            yoy_t1 = (rs[-2] - rs[-3]) / abs(rs[-3]) if len(rs) >= 3 and rs[-3] else np.nan
            out[code] = (yoy_t - yoy_t1) if yoy_t1 == yoy_t1 else np.nan
        else:
            out[code] = np.nan
    return pd.DataFrame({'FCF_加速': out})


# ─────────────── 2. 总分_delta_2y (company_annual_scores PIT 内存切片) ───────────────
def load_total_score_delta2y(keys):
    """keys=[(code,date)]。用 _SCORE_CACHE 全历史, PIT 切 report_year≤pit_max 取 T-(T-2)。"""
    if not keys:
        return pd.DataFrame(columns=['总分_delta_2y'])
    out = {}
    for code, date in keys:
        g = _SCORE_CACHE.get(code)
        if g is None or g.empty:
            out[code] = np.nan; continue
        sub = g[g['report_year'] <= _pit_max(date)]
        rs = sub['total_score'].tail(3).values
        out[code] = (rs[-1] - rs[-3]) if len(rs) >= 3 else np.nan
    return pd.DataFrame({'总分_delta_2y': out})


# ─────────────── 3. nb_hold_ratio (pro.hk_hold ≤ date) ───────────────
_HK_CACHE = {}


def load_nb_hold(codes, date_yyyymmdd):
    """pro.hk_hold 全量缓存/股, 截 trade_date≤date: nb_hold_ratio(最新) + nb_hold_chg_20d/60d
    (与 20/60 交易日前 ratio 之差, PIT)。恢复历史 1w 模型用的 nb_hold_chg_20d/60d(原特征名)。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    pro = ts.pro_api()
    d = int(str(date_yyyymmdd))
    ratio_out = {}; chg20 = {}; chg60 = {}
    for c in codes:
        if c not in _HK_CACHE:
            try:
                df = pro.hk_hold(ts_code=c, fields='trade_date,ratio')
                _HK_CACHE[c] = df
            except Exception:
                _HK_CACHE[c] = pd.DataFrame()
        df = _HK_CACHE[c]
        if df is None or df.empty:
            ratio_out[c] = chg20[c] = chg60[c] = np.nan; continue
        sub = df[pd.to_numeric(df['trade_date'], errors='coerce') <= d]
        if sub.empty:
            ratio_out[c] = chg20[c] = chg60[c] = np.nan; continue
        r = float(sub['ratio'].iloc[-1]); ratio_out[c] = r
        vals = sub['ratio'].astype(float).to_numpy()
        chg20[c] = float(r - vals[-21]) if len(vals) >= 21 else np.nan   # 20 交易日前
        chg60[c] = float(r - vals[-61]) if len(vals) >= 61 else np.nan   # 60 交易日前
    return pd.DataFrame({'nb_hold_ratio': pd.Series(ratio_out),
                         'nb_hold_chg_20d': pd.Series(chg20),
                         'nb_hold_chg_60d': pd.Series(chg60)})


# ─────────────── 4. PB_vs同行中位 (industry_daily + daily_basic PIT, 规范口径) ───────────────
_PB_STOCK_CACHE = {}    # code → daily_basic[trade_date, pb] 全量(跨截面复用, 每股只取 1 次 API)
_INDPB_CACHE = None     # (s2i 映射, industry_daily 全量 DataFrame) 一次查, 内存切片


def load_pb_vs_industry(keys):
    """keys=[(code, date)]。个股 daily_basic PB / 行业 industry_daily sw_index_pb, 都 PIT ≤date 取最近。
    返回 {(str(code), str(date)): 个股PB/行业PB}。

    **规范口径**(2026-06 定): 生产 SC 模型 PB_vs_同行中位 统一采此 PIT 行业口径, 与回测同源。
    个股 daily_basic 每股全量缓存(定增/回测跨多截面时每股只 1 次 API); industry_daily 全量一次查。

    **向量化 PIT**(2026-06-26): 缓存预算每股/每行业的 (td_int 排序, pb) 数组, 按 searchsorted 取 ≤date 最新;
    旧实现逐 key 调 pd.to_numeric(全量 trade_date) + 逐日期全量 filter 1.1M 行 → 全A 563k 对卡 46min。语义不变。"""
    if not keys:
        return {}
    pro = _pro()
    global _INDPB_CACHE
    if _INDPB_CACHE is None:
        conn = _conn()
        idmap = pd.read_sql('SELECT stock_code, index_code FROM industry_data', conn)
        idf = pd.read_sql("SELECT index_code, trade_date, pb FROM industry_daily", conn)
        conn.close()
        idf['td_int'] = pd.to_numeric(idf['trade_date'], errors='coerce')
        idf = idf.dropna(subset=['td_int', 'pb'])             # 行业 .last() skipna 语义: 弃 NaN pb 行
        idf = idf.sort_values(['index_code', 'td_int'])
        ind_arrays = {str(ic): (g['td_int'].to_numpy(np.int64), g['pb'].to_numpy(float))
                      for ic, g in idf.groupby('index_code')}   # 每 index: (排序 td_int, pb)
        _INDPB_CACHE = (dict(zip(idmap['stock_code'].astype(str), idmap['index_code'].astype(str))), ind_arrays)
    s2i, ind_arrays = _INDPB_CACHE

    # 本地优先: 批量读 stock_daily_basic.pb 填 _PB_STOCK_CACHE(存每股 (td_int 排序, pb) 数组)
    # 分块读(全A 6M 行一次性 read_sql 峰值~1.5GB; 每块800股, 峰值内存有界)
    need = [str(c) for c, _ in keys if str(c) not in _PB_STOCK_CACHE]
    if need:
        try:
            import pymysql
            from utils.db_manager import ValuationDB
            _cfg = ValuationDB.MYSQL_CONFIG
            _cn = pymysql.connect(host=_cfg['host'], port=_cfg['port'], user=_cfg['user'],
                                  password=_cfg['password'], database=_cfg['database'], charset=_cfg['charset'])
            try:
                for _i in range(0, len(need), 800):
                    _blk = need[_i:_i + 800]
                    _ph = ','.join(['%s'] * len(_blk))
                    _dfb = pd.read_sql(f"SELECT stock_code,trade_date,pb FROM stock_daily_basic WHERE stock_code IN ({_ph})",
                                       _cn, params=_blk)
                    for _c, _g in _dfb.groupby('stock_code'):
                        _g = _g.copy(); _g['td_int'] = pd.to_numeric(_g['trade_date'], errors='coerce')
                        _g = _g.dropna(subset=['td_int']).sort_values('td_int')
                        _PB_STOCK_CACHE[str(_c)] = (_g['td_int'].to_numpy(np.int64), _g['pb'].to_numpy(float))
            finally:
                _cn.close()
        except Exception:
            pass

    out = {}
    for code, date in keys:
        code, d = str(code), str(date)
        try:
            d_int = int(float(d))
        except ValueError:
            d_int = 0
        # 个股 PB: searchsorted ≤ d_int 最新(等价旧 sdf[td≤d].iloc[-1])
        if code not in _PB_STOCK_CACHE:
            try:
                _d = pro.daily_basic(ts_code=code, fields='trade_date,pb')
                _d['td_int'] = pd.to_numeric(_d['trade_date'], errors='coerce')
                _d = _d.dropna(subset=['td_int']).sort_values('td_int')
                _PB_STOCK_CACHE[code] = (_d['td_int'].to_numpy(np.int64), _d['pb'].to_numpy(float))
            except Exception:
                _PB_STOCK_CACHE[code] = (np.array([], np.int64), np.array([], float))
        td_arr, pb_arr = _PB_STOCK_CACHE[code]
        stock_pb = np.nan
        if len(td_arr):
            pos = int(np.searchsorted(td_arr, d_int, 'right')) - 1
            if pos >= 0:
                stock_pb = float(pb_arr[pos])
        # 行业 PB: searchsorted ≤ d_int 最新(等价旧 groupby.last(skipna))
        ic = s2i.get(code); ipb = np.nan
        if ic and d_int and ic in ind_arrays:
            itd, ipb_arr = ind_arrays[ic]
            if len(itd):
                ipos = int(np.searchsorted(itd, d_int, 'right')) - 1
                if ipos >= 0:
                    ipb = float(ipb_arr[ipos])
        out[(code, d)] = (stock_pb / ipb) if stock_pb == stock_pb and ipb == ipb and ipb else np.nan
    return out


# ─────────────── 5. sue_beat (forecast/express/income 披露 PIT) ───────────────
def load_sue_beat(keys):
    """SUE 全系列(2026-06-26 扩展, 原只出 sue_beat): 输出 sue_beat + sue_zscore + sue_pos_streak +
    sue_recency_d + sue_yoy_acc + sue_yoy_mean3 + sue_yoy + sue_up_trend。PIT: ann_date≤date 最近披露。
    复用 fetch_factors._sue_for_sample(已算全变体); 需先 prefetch_sue_timelines, 否则惰性逐股建(慢)。
    keys 单截面用(codes, 同一 date)→ 输出 index=code。原特征名, 现有 1w/2w 模型直接消费。"""
    from data_pipeline.fetch_factors import _sue_for_sample, _build_disclosure_timeline
    SUE_COLS = ['sue_beat', 'sue_zscore', 'sue_pos_streak', 'sue_recency_d',
                'sue_yoy_acc', 'sue_yoy_mean3', 'sue_yoy', 'sue_up_trend']
    if not keys:
        return pd.DataFrame(columns=SUE_COLS)
    pro = None
    rows = {}
    for code, date in keys:
        c = str(code)
        if c not in _SUE_CACHE:                  # 未预取 → 惰性建(逐股 3 API; 批量请先 prefetch)
            pro = pro or _pro()
            try:
                _SUE_CACHE[c] = _build_disclosure_timeline(pro, c)
            except Exception:
                _SUE_CACHE[c] = None
        tl = _SUE_CACHE[c]
        d = _sue_for_sample(tl, str(date)) if tl is not None else {}
        rows[c] = {col: d.get(col, np.nan) for col in SUE_COLS}
    return pd.DataFrame.from_dict(rows, orient='index', columns=SUE_COLS)


def load_specials(codes, date_yyyymmdd):
    """一次性加载 5 个特殊特征, 返回 DataFrame(index=codes)。"""
    d = str(date_yyyymmdd)
    keys = [(c, d) for c in codes]
    pb_map = load_pb_vs_industry(keys)   # {(code, date): ratio}
    pb_series = pd.Series({c: pb_map.get((c, d), np.nan) for c in codes}, name='PB_vs_同行中位')
    frames = [
        load_fcf_accel(keys),
        load_total_score_delta2y(keys),
        load_nb_hold(codes, date_yyyymmdd),
        pd.DataFrame({'PB_vs_同行中位': pb_series}),
        load_sue_beat(keys),
    ]
    df = pd.concat(frames, axis=1)
    df.index.name = 'code'
    return df
