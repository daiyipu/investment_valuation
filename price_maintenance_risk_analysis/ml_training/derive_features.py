#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增特征衍生脚本 - 从基础特征派生高级特征

两阶段流水线:
  Stage 1 (A-E): 纯 parquet 运算，不访问数据库
  Stage 2 (F-G): 需 MySQL 查询 (industry_daily, market_indices)

用法:
    python ml_training/derive_features.py [input.parquet] [--output features_derived.parquet] [--no-db]

输出:
    ml_training/data/features_derived.parquet - 原始特征 + 衍生特征
"""

import sys
import os
import argparse
import numpy as np
import pandas as pd


# ====== 通用工具 ======

def _safe_divide(a, b, fillna=np.nan):
    """安全除法: abs(b)<1e-8 返回 NaN"""
    a = pd.to_numeric(a, errors='coerce')
    b = pd.to_numeric(b, errors='coerce')
    result = np.where(np.abs(b) < 1e-8, fillna, a / b)
    return pd.Series(result, index=a.index)


def _safe_yoy(current, prior):
    """YoY增长率: (current - prior) / |prior|"""
    current = pd.to_numeric(current, errors='coerce')
    prior = pd.to_numeric(prior, errors='coerce')
    return _safe_divide(current - prior, prior)


def _check_cols(df, required, category):
    """检查必需列是否存在，缺失则返回False"""
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f'  ⚠️ {category}: 缺少列 {missing[:5]}{"..." if len(missing)>5 else ""}，跳过')
        return False
    return True


# ====== A类: FCF 增长率 (+18特征) ======

FCF_METRICS = ['营收', 'NOPAT', 'FCF', '营业利润', '净利润_fcf', '资本支出']


def derive_fcf_growth_rates(df):
    """从 _T/_T1/_T2 列计算 YoY、CAGR2、增长加速度"""
    print('\n  A类: FCF增长率...')
    required = []
    for m in FCF_METRICS:
        for s in ['_T', '_T1', '_T2']:
            required.append(f'{m}{s}')
    if not _check_cols(df, required, 'A类-FCF增长率'):
        return df

    new_cols = {}
    for m in FCF_METRICS:
        t  = pd.to_numeric(df[f'{m}_T'],  errors='coerce')
        t1 = pd.to_numeric(df[f'{m}_T1'], errors='coerce')
        t2 = pd.to_numeric(df[f'{m}_T2'], errors='coerce')

        # YoY增长率
        new_cols[f'{m}_YoY'] = _safe_yoy(t, t1)

        # 2年CAGR: (T/T2)^0.5 - 1 (仅当同号时)
        ratio = _safe_divide(t, t2)
        same_sign = (t * t2) > 0
        cagr = np.where(same_sign, np.power(ratio.abs(), 0.5) - 1, np.nan)
        cagr = np.where(same_sign & (t2 > 0), ratio.pow(0.5) - 1, cagr)
        new_cols[f'{m}_CAGR2'] = pd.Series(cagr, index=df.index)

        # 增长加速度: YoY_T - YoY_T1
        yoy_t  = _safe_yoy(t, t1)
        yoy_t1 = _safe_yoy(t1, t2)
        new_cols[f'{m}_加速'] = yoy_t - yoy_t1

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in list(coverage.items())[:6])}...')
    return df


# ====== B类: FCF 交叉比率 (+5特征) ======

def derive_fcf_cross_metrics(df):
    """FCF相关的财务比率"""
    print('\n  B类: FCF交叉比率...')
    required = ['FCF_T', '营收_T', 'NOPAT_T', '资本支出_T', '折旧_T']
    if not _check_cols(df, required, 'B类-FCF交叉比率'):
        return df

    fcf    = pd.to_numeric(df['FCF_T'],    errors='coerce')
    rev    = pd.to_numeric(df['营收_T'],   errors='coerce')
    nopat  = pd.to_numeric(df['NOPAT_T'],  errors='coerce')
    capex  = pd.to_numeric(df['资本支出_T'], errors='coerce')
    dep    = pd.to_numeric(df['折旧_T'],   errors='coerce')

    new_cols = {
        'FCF_margin':     _safe_divide(fcf, rev),
        'FCF_conversion': _safe_divide(fcf, nopat),
        'capex_to_dep':   _safe_divide(capex, dep),
        'capex_intensity': _safe_divide(capex, rev),
        'NOPAT_margin':   _safe_divide(nopat, rev),
    }
    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


# ====== C类: 财务评分变动 (+9特征) ======

SCORE_METRICS = ['总分', '盈利能力', '成长能力']
SCORE_DELTAS = [('delta_1y', 'T', 'T-1'), ('delta_2y', 'T', 'T-2'), ('delta_4y', 'T', 'T-4')]


def derive_financial_score_deltas(df):
    """财务评分的年度变动量"""
    print('\n  C类: 财务评分变动...')
    # 检查必需列: {metric}_T 和 {metric}_T-1 等
    required = set()
    for m in SCORE_METRICS:
        for _, y_new, y_old in SCORE_DELTAS:
            required.add(f'{m}_{y_new}')
            required.add(f'{m}_{y_old}')
    if not _check_cols(df, list(required), 'C类-评分变动'):
        return df

    new_cols = {}
    for m in SCORE_METRICS:
        for suffix, y_new, y_old in SCORE_DELTAS:
            col_new = f'{m}_{y_new}'
            col_old = f'{m}_{y_old}'
            if col_new in df.columns and col_old in df.columns:
                v_new = pd.to_numeric(df[col_new], errors='coerce')
                v_old = pd.to_numeric(df[col_old], errors='coerce')
                new_cols[f'{m}_{suffix}'] = v_new - v_old

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in list(coverage.items())[:6])}...')
    return df


# ====== D类: 估值相对特征 (+7特征) ======

def derive_valuation_relative(df):
    """个股估值 vs 行业/同行"""
    print('\n  D类: 估值相对特征...')

    new_cols = {}

    # vs 行业
    pairs_industry = [('PE', '个股PE', '行业PE'), ('PB', '个股PB', '行业PB')]
    for val_type, col_stock, col_ind in pairs_industry:
        if col_stock in df.columns and col_ind in df.columns:
            new_cols[f'{val_type}_vs_行业'] = _safe_divide(
                pd.to_numeric(df[col_stock], errors='coerce'),
                pd.to_numeric(df[col_ind], errors='coerce')
            )

    # vs 同行
    pairs_peer = [
        ('PE', '个股PE', '同行PE_均值', 'PE_vs_同行均值'),
        ('PE', '个股PE', '同行PE_中位', 'PE_vs_同行中位'),
        ('PB', '个股PB', '同行PB_均值', 'PB_vs_同行均值'),
        ('PB', '个股PB', '同行PB_中位', 'PB_vs_同行中位'),
        ('PS', '个股PS', '同行PS_均值', 'PS_vs_同行均值'),
    ]
    for val_type, col_stock, col_peer, col_name in pairs_peer:
        if col_stock in df.columns and col_peer in df.columns:
            new_cols[col_name] = _safe_divide(
                pd.to_numeric(df[col_stock], errors='coerce'),
                pd.to_numeric(df[col_peer], errors='coerce')
            )

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


def derive_pb_vs_industry_pit(df):
    """PB_vs_同行中位 ← PIT 行业口径(规范, 2026-06 定): 覆盖 derive_valuation_relative 的 peer 中位口径。

    原 peer 口径 = 个股PB / peer_companies 同行中位(非 PIT 快照, 覆盖~53%, 且是潜在 PIT 泄漏)。
    新口径 = daily_basic PB(≤报价日) / industry_daily sw_index_pb(≤报价日), 与回测 feature_loaders 同源。
    生产 SC 模型 15 特征含此列 → 重训后 PB WOE 基于此口径, 回测/定增同分布不再错配。

    前置: 定增 universe 的 industry_daily 须全历史重摄(`ingest_raw --universe placement --industry-only`),
          否则早年 issue_date 的行业 PB 缺失 → 覆盖偏低。
    """
    print('\n  D类补: PB_vs_同行中位 ← PIT 行业口径(覆盖 peer 中位)')
    if '股票代码' not in df.columns or '报价日' not in df.columns:
        print('    缺 股票代码/报价日, 跳过(保留原 peer 口径)')
        return df
    pkg = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if pkg not in sys.path:
        sys.path.insert(0, pkg)
    sys.path.insert(0, os.path.join(pkg, 'scripts'))
    from feature_loaders import load_pb_vs_industry

    pairs = [(str(c), str(d)) for c, d in zip(df['股票代码'], df['报价日']) if pd.notna(d)]
    pb_map = load_pb_vs_industry(pairs)
    vals = [pb_map.get((str(c), str(d)), np.nan)
            for c, d in zip(df['股票代码'], df['报价日'])]
    df['PB_vs_同行中位'] = vals
    cov = pd.Series(vals).notna().mean() * 100
    print(f'    PB_vs_同行中位(PIT 行业口径): 覆盖 {cov:.1f}%')
    if cov < 50:
        print('    ⚠️ 覆盖偏低: 多半是 industry_daily 未全历史重摄 → 先跑 '
              '`ingest_raw --universe placement --industry-only` 再重训')
    return df


# ====== E类: 行情动量与位置 (+7特征) ======

def derive_market_momentum(df):
    """波动率比值、动量加速、价格vs均线"""
    print('\n  E类: 行情动量与位置...')

    new_cols = {}

    # 波动率比值
    for col_name, w_short, w_long in [('vol_ratio_20_60', '波动率_20d', '波动率_60d'),
                                       ('vol_ratio_60_250', '波动率_60d', '波动率_250d')]:
        if w_short in df.columns and w_long in df.columns:
            new_cols[col_name] = _safe_divide(
                pd.to_numeric(df[w_short], errors='coerce'),
                pd.to_numeric(df[w_long], errors='coerce')
            )

    # 动量加速
    if '年化收益_60d' in df.columns and '年化收益_120d' in df.columns:
        new_cols['return_acceleration'] = (
            pd.to_numeric(df['年化收益_60d'], errors='coerce') -
            pd.to_numeric(df['年化收益_120d'], errors='coerce')
        )

    # 价格 vs 均线
    if '当前价' in df.columns:
        price = pd.to_numeric(df['当前价'], errors='coerce')
        for ma_label in ['MA20', 'MA60', 'MA120', 'MA250']:
            if ma_label in df.columns:
                new_cols[f'price_vs_{ma_label}'] = _safe_divide(price, pd.to_numeric(df[ma_label], errors='coerce'))

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in list(coverage.items())[:6])}...')
    return df


# ====== M类: 定增结构衍生(折价率/稀释率/募集市值比/大股东参与/锁定期, 优先级1 domain核心) ======

def derive_placement_structure(df):
    """定增结构衍生特征。原料: 定增_发行价/增发数/募资/股本/发行对象/定价原则/发行方式/解禁日。
    绝对原料(发行价/增发数/募资/股本)不可跨公司比较, 由 feature_exclusions 剔除, 仅留衍生比率/类别。
    数据源: 东方财富 RPT_SEO_DETAIL(81%覆盖) + 易米主表解禁日(20%)。
    """
    print('\n  M类: 定增结构衍生特征...')
    new = {}
    price = pd.to_numeric(df.get('报价日价格'), errors='coerce')
    iprice = pd.to_numeric(df.get('定增_发行价'), errors='coerce')
    inum = pd.to_numeric(df.get('定增_增发数量'), errors='coerce')
    share_b = pd.to_numeric(df.get('定增_发行前股本'), errors='coerce')
    raise_total = pd.to_numeric(df.get('定增_募资总额'), errors='coerce')

    # 折价率 = 发行价/市价 - 1 (负=折价; 折价越深, 解禁抛压越大)
    new['折价率'] = _safe_divide(iprice, price) - 1
    # 稀释率 = 增发数 / 发行前股本
    new['定增稀释率'] = _safe_divide(inum, share_b)
    # 募集市值比 = 募资 / (发行前股本 × 市价) ≈ 募资/发行前总市值
    new['募集市值比'] = _safe_divide(raise_total, share_b * price)

    # 大股东参与(heuristic): 发行对象含 控股/实际控制/实控人, 或与股票简称同名实体
    obj = df.get('定增_发行对象', pd.Series(index=df.index, dtype=object)).fillna('').astype(str)
    name = df.get('股票简称', pd.Series(index=df.index, dtype=object)).fillna('').astype(str)
    kw_hit = obj.str.contains('控股|实际控制|实控人', regex=True)
    name_hit = pd.Series([(len(n) >= 2 and n in o) for n, o in zip(name, obj)], index=df.index)
    big_holder = (kw_hit | name_hit).fillna(False)
    new['定增大股东参与'] = big_holder.astype(int)

    # 锁定期天数: 有解禁日则(解禁日−报价日); 否则规则 18m(大股东/战投)/6m(其他)
    issue = pd.to_datetime(df.get('报价日'), format='%Y%m%d', errors='coerce')
    unlock = pd.to_datetime(df.get('定增_解禁日'), format='%Y%m%d', errors='coerce')
    lock_from_date = (unlock - issue).dt.days
    lock_rule = pd.Series(np.where(big_holder, 540, 180), index=df.index)
    new['定增锁定期天数'] = lock_from_date.combine_first(lock_rule)

    # 询价发行(0/1): 发行方式/定价原则含"询价" → 市场化定价代理
    way = (df.get('定增_发行方式', pd.Series(index=df.index, dtype=object)).fillna('').astype(str)
           + df.get('定增_定价原则', pd.Series(index=df.index, dtype=object)).fillna('').astype(str))
    new['定增_询价发行'] = way.str.contains('询价').astype(int)

    for k, v in new.items():
        df[k] = v
    cov = {k: f'{pd.Series(v).notna().mean()*100:.0f}%' for k, v in new.items()}
    print(f'    +{len(new)}个特征: {", ".join(f"{k}({v})" for k,v in cov.items())}')
    return df


# ====== F类: 行业PE/PB增长率 (+6特征, 需DB) ======

def derive_industry_valuation_growth(df):
    """从 industry_daily 表计算行业PE/PB在报价日前的增长率"""
    print('\n  F类: 行业PE/PB增长率 (需MySQL)...')

    import pymysql
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')

    # 从 industry_data 表获取每只股票对应的 index_code（与 industry_daily 一致）
    stock_codes = df['股票代码'].unique().tolist()
    codes_ph = ','.join(['%s'] * len(stock_codes))
    id_df = pd.read_sql(
        f'SELECT stock_code, index_code FROM industry_data WHERE stock_code IN ({codes_ph})',
        conn, params=stock_codes
    )
    # 建立 stock_code → index_code 映射
    stock_to_index = dict(zip(id_df['stock_code'], id_df['index_code']))
    print(f'    行业映射: {len(stock_to_index)} 只股票')

    # 批量加载所有 industry_daily 数据
    daily_df = pd.read_sql(
        'SELECT index_code, trade_date, pe, pb, close FROM industry_daily ORDER BY index_code, trade_date',
        conn
    )
    conn.close()

    if daily_df.empty:
        print('    ⚠️ industry_daily 无数据，跳过')
        return df

    print(f'    加载行业日线: {daily_df["index_code"].nunique()}个行业, {len(daily_df)}条')

    # 按 index_code 分组
    daily_groups = {}
    for code, group in daily_df.groupby('index_code'):
        g = group.sort_values('trade_date').reset_index(drop=True)
        daily_groups[code] = {
            'dates': pd.to_numeric(g['trade_date'], errors='coerce').fillna(0).astype('int64').values,  # int 比较: trade_date 类型不定(str/int/float), 统一 int64 与报价日数值比
            'pe': g['pe'].values.astype(float),
            'pb': g['pb'].values.astype(float),
        }

    # 获取每条样本的 index_code
    sample_index_codes = df['股票代码'].map(stock_to_index)

    # 对每条样本计算
    growth_windows = [60, 120, 250]
    pe_growth = {w: np.full(len(df), np.nan) for w in growth_windows}
    pb_growth = {w: np.full(len(df), np.nan) for w in growth_windows}

    matched = 0
    for i, (idx, row) in enumerate(df.iterrows()):   # 用位置 i(非索引值 idx): df.index 可能是股票代码(str), iloc/growth 数组须按位置
        code = sample_index_codes.iloc[i] if i < len(sample_index_codes) else None
        if pd.isna(code) or code not in daily_groups:
            continue
        issue_date_raw = row.get('报价日')
        if pd.isna(issue_date_raw):
            continue
        try:
            issue_int = int(float(issue_date_raw))
        except (ValueError, TypeError):
            continue
        if issue_int < 10101:
            continue

        dg = daily_groups[code]
        # 找到 <= issue_int 的最大索引(int 比较, 免 str/int 类型冲突)
        mask = dg['dates'] <= issue_int
        if not mask.any():
            continue
        pos = mask.sum() - 1  # 最后一个 <= issue_str 的位置

        pe_now = dg['pe'][pos]
        pb_now = dg['pb'][pos]
        if np.isnan(pe_now) and np.isnan(pb_now):
            continue

        for w in growth_windows:
            if pos >= w:
                pe_prev = dg['pe'][pos - w]
                pb_prev = dg['pb'][pos - w]
                if not np.isnan(pe_now) and not np.isnan(pe_prev) and abs(pe_prev) > 1e-8:
                    pe_growth[w][i] = pe_now / pe_prev - 1
                if not np.isnan(pb_now) and not np.isnan(pb_prev) and abs(pb_prev) > 1e-8:
                    pb_growth[w][i] = pb_now / pb_prev - 1
        matched += 1

    print(f'    匹配行业日线: {matched}/{len(df)}')

    new_cols = {}
    for w in growth_windows:
        new_cols[f'行业PE_{w}d增长'] = pd.Series(pe_growth[w], index=df.index)
        new_cols[f'行业PB_{w}d增长'] = pd.Series(pb_growth[w], index=df.index)

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


# ====== G类: 大盘指数特征 (+6特征, 需DB) ======

# 交易所 → 指数名 映射
EXCHANGE_INDEX_MAP = {
    '300': '创业板指',
    '301': '创业板指',
    '688': '科创50',
    '60':  '沪深300',
    '000': '中证500',
    '001': '中证500',
    '002': '中证500',
    '003': '中证500',
}
FALLBACK_INDEX = '沪深300'


def _get_index_name(stock_code):
    """根据股票代码前缀判断对应的市场指数"""
    for prefix, idx_name in sorted(EXCHANGE_INDEX_MAP.items(), key=lambda x: -len(x[0])):
        if stock_code.startswith(prefix):
            return idx_name
    return FALLBACK_INDEX


def derive_market_index_features(df):
    """从 market_indices 表获取报价日时的大盘特征"""
    print('\n  G类: 大盘指数特征 (需MySQL)...')

    import pymysql
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')

    # 批量加载所有数据
    query = '''SELECT index_name, locked_date, current_level,
                      volatility_60d, volatility_120d, volatility_250d,
                      return_60d, return_120d,
                      ma_250, win_rate_120d
               FROM market_indices'''
    mi_df = pd.read_sql(query, conn)
    conn.close()

    if mi_df.empty:
        print('    ⚠️ market_indices 无数据，跳过')
        return df

    print(f'    加载大盘指数: {mi_df["index_name"].nunique()}个指数, {len(mi_df)}条')

    # 构建 (index_name, locked_date) → row 的查找字典
    mi_lookup = {}
    for _, row in mi_df.iterrows():
        key = (row['index_name'], str(row['locked_date']))
        mi_lookup[key] = row

    # 对每条样本查找
    mkt_vol120 = np.full(len(df), np.nan)
    mkt_ret120 = np.full(len(df), np.nan)
    mkt_wr120  = np.full(len(df), np.nan)
    mkt_above_ma250 = np.full(len(df), np.nan)
    mkt_dist_ma250  = np.full(len(df), np.nan)
    mkt_vol_ratio   = np.full(len(df), np.nan)

    for idx, row in df.iterrows():
        stock_code = str(row.get('股票代码', ''))
        issue_date_raw = row.get('报价日')
        if pd.isna(issue_date_raw):
            continue
        issue_str = str(int(float(issue_date_raw))) if isinstance(issue_date_raw, (int, float)) else str(issue_date_raw)
        if len(issue_str) < 8:
            continue

        index_name = _get_index_name(stock_code)

        # 尝试精确匹配，找不到则往前找最近的
        mi_row = None
        for key in [(index_name, issue_str)]:
            if key in mi_lookup:
                mi_row = mi_lookup[key]
                break

        if mi_row is None:
            # 查找该指数在该日期之前的最近日期
            candidates = mi_df[(mi_df['index_name'] == index_name) &
                              (mi_df['locked_date'] <= issue_str)]
            if candidates.empty:
                # 回退到沪深300
                candidates = mi_df[(mi_df['index_name'] == FALLBACK_INDEX) &
                                  (mi_df['locked_date'] <= issue_str)]
            if candidates.empty:
                continue
            mi_row = candidates.iloc[-1]

        level = mi_row.get('current_level')
        vol120 = mi_row.get('volatility_120d')
        vol60  = mi_row.get('volatility_60d')
        vol250 = mi_row.get('volatility_250d')
        ret120 = mi_row.get('return_120d')
        wr120  = mi_row.get('win_rate_120d')
        ma250  = mi_row.get('ma_250')

        if pd.notna(vol120):
            mkt_vol120[idx] = float(vol120)
        if pd.notna(ret120):
            mkt_ret120[idx] = float(ret120)
        if pd.notna(wr120):
            mkt_wr120[idx] = float(wr120)
        if pd.notna(level) and pd.notna(ma250) and float(ma250) > 0:
            mkt_above_ma250[idx] = 1.0 if float(level) > float(ma250) else 0.0
            mkt_dist_ma250[idx] = (float(level) - float(ma250)) / float(ma250)
        if pd.notna(vol60) and pd.notna(vol250) and abs(float(vol250)) > 1e-8:
            mkt_vol_ratio[idx] = float(vol60) / float(vol250)

    new_cols = {
        '市场波动率_120d':   pd.Series(mkt_vol120, index=df.index),
        '市场年化收益_120d': pd.Series(mkt_ret120, index=df.index),
        '市场胜率_120d':     pd.Series(mkt_wr120, index=df.index),
        '市场_above_MA250':  pd.Series(mkt_above_ma250, index=df.index),
        '市场距离MA250':     pd.Series(mkt_dist_ma250, index=df.index),
        '市场波动率比值':    pd.Series(mkt_vol_ratio, index=df.index),
    }

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


# ====== 模块级 OHLCV 缓存(策略信号 + 因子引擎共用, 避免重复 pro_bar) ======
_OHLCV_CACHE = {}   # code -> (dates, {open/high/low/close/vol/amount} np arrays) 或 (None, None)


def _fetch_ohlcv_raw(code):
    """实际调 pro_bar(无缓存), 返回 (dates, ohlcv_dict); 失败抛异常。供并发预取重试用。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    d = ts.pro_bar(ts_code=code, adj='qfq')
    if d is None or len(d) == 0:
        raise ValueError(f'pro_bar 空返回: {code}')
    d = d.sort_values('trade_date').drop_duplicates('trade_date').reset_index(drop=True)
    amt = d['amount'].astype(float).values if 'amount' in d.columns else d['vol'].astype(float).values
    return (d['trade_date'].values, {
        'open': d['open'].astype(float).values, 'high': d['high'].astype(float).values,
        'low': d['low'].astype(float).values, 'close': d['close'].astype(float).values,
        'vol': d['vol'].astype(float).values, 'amount': amt,
    })


def _get_stock_ohlcv(code):
    """pro_bar qfq 全量(缓存), 返回 (dates, ohlcv_dict) 或 (None, None)。
    供 derive_strategy_signals(取 close) 与 derive_alpha_beta_factors(OHLCV) 共用。"""
    if code in _OHLCV_CACHE:
        return _OHLCV_CACHE[code]
    try:
        res = _fetch_ohlcv_raw(code)
    except Exception:
        res = (None, None)
    _OHLCV_CACHE[code] = res
    return res


def prefetch_ohlcv(codes, max_workers=12):
    """并发预取所有 unique 股票 OHLCV 入缓存(I/O bound, 线程即可)。
    tushare 限流时单股退避重试(3 次)。已缓存的跳过。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import time
    todo = [c for c in dict.fromkeys(str(c) for c in codes) if c not in _OHLCV_CACHE]
    if not todo:
        print(f'    OHLCV 缓存已热({len(_OHLCV_CACHE)}), 跳过预取')
        return
    print(f'    并发预取 {len(todo)} 只 OHLCV (max_workers={max_workers}, 限流重试)...')

    def _work(code):
        for attempt in range(3):
            try:
                _OHLCV_CACHE[code] = _fetch_ohlcv_raw(code)
                return True
            except Exception:
                time.sleep(1.5 * (attempt + 1))   # 限流/网络 → 退避重试
        _OHLCV_CACHE[code] = (None, None)
        return False

    ok = 0
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = [ex.submit(_work, c) for c in todo]
        for n, fut in enumerate(as_completed(futs), 1):
            if fut.result():
                ok += 1
            if n % 50 == 0:
                print(f'    预取 {n}/{len(todo)} (成功 {ok})')
    print(f'    预取完成: {ok}/{len(todo)} 成功')


# ====== 模块级 daily_basic 缓存(换手率/量比, 与 OHLCV 缓存并列) ======
_DAILY_BASIC_CACHE = {}   # code -> (dates, {turnover, vol_ratio}) 或 (None, None)


def _fetch_daily_basic_raw(code):
    """实际调 daily_basic(无缓存), 返回 (dates, {turnover, vol_ratio}); 失败抛异常。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    d = ts.pro_api().daily_basic(ts_code=code,
                                 fields='ts_code,trade_date,turnover_rate,volume_ratio')
    if d is None or len(d) == 0:
        raise ValueError(f'daily_basic 空返回: {code}')
    d = d.sort_values('trade_date').drop_duplicates('trade_date').reset_index(drop=True)
    return (d['trade_date'].values,
            {'turnover': d['turnover_rate'].astype(float).values,
             'vol_ratio': d['volume_ratio'].astype(float).values})


def _get_daily_basic(code):
    """daily_basic 全量(缓存), 返回 (dates, {turnover, vol_ratio}) 或 (None, None)。"""
    if code in _DAILY_BASIC_CACHE:
        return _DAILY_BASIC_CACHE[code]
    try:
        res = _fetch_daily_basic_raw(code)
    except Exception:
        res = (None, None)
    _DAILY_BASIC_CACHE[code] = res
    return res


def prefetch_daily_basic(codes, max_workers=12):
    """并发预取所有 unique 股票 daily_basic 入缓存。tushare 限流时单股退避重试(3 次)。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import time
    todo = [c for c in dict.fromkeys(str(c) for c in codes) if c not in _DAILY_BASIC_CACHE]
    if not todo:
        print(f'    daily_basic 缓存已热({len(_DAILY_BASIC_CACHE)}), 跳过预取')
        return
    print(f'    并发预取 {len(todo)} 只 daily_basic (max_workers={max_workers}, 限流重试)...')

    def _work(code):
        for attempt in range(3):
            try:
                _DAILY_BASIC_CACHE[code] = _fetch_daily_basic_raw(code)
                return True
            except Exception:
                time.sleep(1.5 * (attempt + 1))
        _DAILY_BASIC_CACHE[code] = (None, None)
        return False

    ok = 0
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = [ex.submit(_work, c) for c in todo]
        for n, fut in enumerate(as_completed(futs), 1):
            if fut.result():
                ok += 1
            if n % 50 == 0:
                print(f'    预取 {n}/{len(todo)} (成功 {ok})')
    print(f'    预取完成: {ok}/{len(todo)} 成功')


# ====== H/I类: 三浪/抵抗策略信号 (需 tushare 网络) ======

def derive_strategy_signals(df):
    """按每行报价日 PIT 回算 三浪/抵抗 信号, 写 5 个特征列。

    bulk-load 优化(避免逐样本打 tushare):
      - 大盘: index_daily('000300.SH') 全量 1 次
      - 个股: 每 unique 股票 pro_bar qfq 全量 1 次(缓存), 切片 ≤报价日
      - 行业: industry_daily 全量 1 次 MySQL(同F类), 切片 ≤报价日
    PIT 由 ≤报价日 内存切片保证。单股失败填 NaN, 不影响整体。
    """
    print('\n  H/I类: 三浪/抵抗策略信号 (需tushare网络)...')
    import time
    pkg = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # price_maintenance_risk_analysis
    if pkg not in sys.path:
        sys.path.insert(0, pkg)
    from strategies.data_loader import _align_three, _MARKET_INDEX
    from strategies.wave3 import wave3_signal
    from strategies.resist import resist_score

    import pymysql
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')

    # 1) stock → industry index_code 映射(实测 industry_daily 键 = industry_data.index_code)
    id_map_df = pd.read_sql('SELECT stock_code, index_code FROM industry_data', conn)
    stock_to_idx = dict(zip(id_map_df['stock_code'], id_map_df['index_code']))

    # 2) 行业日线全量(按 index_code 分组, 升序)
    ind_df = pd.read_sql('SELECT index_code, trade_date, close FROM industry_daily ORDER BY index_code, trade_date', conn)
    conn.close()
    ind_groups = {}
    if not ind_df.empty:
        for code, g in ind_df.groupby('index_code'):
            g = g.sort_values('trade_date').reset_index(drop=True)
            ind_groups[code] = (g['trade_date'].values, g['close'].astype(float).values)
    print(f'    行业日线: {len(ind_groups)} 个行业')

    # 3) 大盘全量(1 次 tushare)
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    mrt = ts.pro_api().index_daily(ts_code=_MARKET_INDEX)
    mrt = mrt.sort_values('trade_date').reset_index(drop=True)
    mkt_dates = mrt['trade_date'].values
    mkt_close = mrt['close'].astype(float).values
    print(f'    大盘 {_MARKET_INDEX}: {len(mrt)} 条日线')

    # 4) 个股 qfq 全量缓存(每 unique 股票 1 次)
    unique_codes = df['股票代码'].astype(str).unique().tolist()
    print(f'    取个股 qfq 全量: {len(unique_codes)} 只(缓存)...')
    prefetch_ohlcv(unique_codes)   # 并发预取(下次跑也复用); 之后 get_stock_close 直接读缓存

    def get_stock_close(code):
        """委托模块级 _get_stock_ohlcv(与因子引擎共用缓存), 取 close。"""
        dates, ohlcv = _get_stock_ohlcv(code)
        if dates is None:
            return (None, None)
        return (dates, ohlcv['close'])

    def _slice(dates, close, end_str, maxlen=300):
        """≤end_str 末段 maxlen。(dates, close) 升序。"""
        if dates is None:
            return None, None
        mask = dates <= end_str
        if not mask.any():
            return None, None
        n = mask.sum()
        lo = max(0, n - maxlen)
        return dates[mask][lo:], close[mask][lo:]

    # 5) 逐样本回算
    maxlen = 300
    w3_score = np.full(len(df), np.nan)
    w3_gain = np.full(len(df), np.nan)
    w3_retr = np.full(len(df), np.nan)
    rs_score = np.full(len(df), np.nan)
    rs_div = np.full(len(df), np.nan)
    fetched = 0

    for i, row in enumerate(df.itertuples(index=True)):
        code = str(getattr(row, '股票代码', ''))
        d_raw = getattr(row, '报价日', None)
        if d_raw is None or pd.isna(d_raw):
            continue
        d_str = str(int(float(d_raw))) if isinstance(d_raw, (int, float, np.floating)) else str(d_raw)
        if len(d_str) < 8:
            continue
        try:
            sd, sc = get_stock_close(code)
            sd2, sc2 = _slice(sd, sc, d_str, maxlen)
            if sd2 is None or len(sc2) < 250:
                continue
            # 三浪
            w = wave3_signal(sc2)
            w3_gain[i] = w['gain'] if w['gain'] else np.nan
            w3_retr[i] = w['retr'] if w['retr'] else np.nan
            w3_score[i] = w['score'] if w['trigger'] else 0.0
            # 抵抗(个股/行业/大盘 对齐收益)
            idx = stock_to_idx.get(code)
            ig = ind_groups.get(idx)
            id2 = cd2 = None
            if ig is not None:
                id2, cd2 = _slice(ig[0], ig[1], d_str, maxlen)
            md2, mc2 = _slice(mkt_dates, mkt_close, d_str, maxlen)
            if id2 is not None and md2 is not None:
                sdf = pd.DataFrame({'trade_date': sd2[-len(id2):], 'close': sc2[-len(id2):]})
                idf = pd.DataFrame({'trade_date': id2, 'close': cd2})
                mdf = pd.DataFrame({'trade_date': md2, 'close': mc2})
                sr, kr, mr = _align_three(sdf, idf, mdf, maxlen)
                if len(sr) >= 65:
                    r = resist_score(sr, kr, mr)
                    rs_score[i] = r['score']
                    rs_div[i] = r['corr_div_stock'] if r['corr_div_stock'] is not None else np.nan
            fetched += 1
        except Exception as e:
            continue
        if (i + 1) % 100 == 0:
            print(f'    进度 {i+1}/{len(df)} | 有信号 {fetched}')

    new_cols = {
        '三浪_score': pd.Series(w3_score, index=df.index),
        '三浪_gain':  pd.Series(w3_gain, index=df.index),
        '三浪_retr':  pd.Series(w3_retr, index=df.index),
        '抵抗_score': pd.Series(rs_score, index=df.index),
        '抵抗_corr_div_stock': pd.Series(rs_div, index=df.index),
    }
    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    匹配 {fetched}/{len(df)} | +{len(new_cols)}特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


# ====== L类: Beta + Alpha158 缺失族因子 (factor_engine, 需 tushare 网络) ======

def derive_alpha_beta_factors(df):
    """按每行报价日 PIT 回算 Beta + Alpha158 缺失族因子(34 列, factor_engine)。

    复用模块级 _OHLCV_CACHE(策略信号先跑则缓存已热, 不重复 pro_bar); 大盘/行业 1 次取数。
    PIT 由 ≤报价日 内存切片保证。单股/单因子失败填 NaN, 不影响整体。
    """
    print('\n  L类: Beta+Alpha158 因子(factor_engine)...')
    pkg = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if pkg not in sys.path:
        sys.path.insert(0, pkg)
    from strategies.data_loader import _align_three, _MARKET_INDEX
    from factor_engine import compute_factors, _ret

    import pymysql
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    id_map_df = pd.read_sql('SELECT stock_code, index_code FROM industry_data', conn)
    stock_to_idx = dict(zip(id_map_df['stock_code'], id_map_df['index_code']))
    ind_df = pd.read_sql('SELECT index_code, trade_date, close FROM industry_daily ORDER BY index_code, trade_date', conn)
    conn.close()
    ind_groups = {}
    if not ind_df.empty:
        for code, g in ind_df.groupby('index_code'):
            g = g.sort_values('trade_date').reset_index(drop=True)
            ind_groups[code] = (g['trade_date'].values, g['close'].astype(float).values)
    print(f'    行业日线: {len(ind_groups)} 个行业')

    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    mrt = ts.pro_api().index_daily(ts_code=_MARKET_INDEX)
    mrt = mrt.sort_values('trade_date').reset_index(drop=True)
    mkt_dates, mkt_close = mrt['trade_date'].values, mrt['close'].astype(float).values

    prefetch_ohlcv(df['股票代码'].astype(str).unique())   # 并发预取(若策略信号先跑则缓存已热, 秒过)
    prefetch_daily_basic(df['股票代码'].astype(str).unique())  # 换手率/量比(量价族 turnover_factors)
    maxlen = 800   # 月线 MACD 需 ~35 个月 ≈ 800 日线(日线技术因子用尾部, 不受影响)
    all_factors = {}
    fetched = 0
    for i, row in enumerate(df.itertuples(index=True)):
        code = str(getattr(row, '股票代码', ''))
        d_raw = getattr(row, '报价日', None)
        if d_raw is None or pd.isna(d_raw):
            continue
        d_str = str(int(float(d_raw))) if isinstance(d_raw, (int, float, np.floating)) else str(d_raw)
        if len(d_str) < 8:
            continue
        try:
            sd, ohlcv = _get_stock_ohlcv(code)
            if sd is None:
                continue
            sm = sd <= d_str
            if not sm.any():
                continue
            n = int(sm.sum()); lo = max(0, n - maxlen)
            o = ohlcv['open'][sm][lo:]; h = ohlcv['high'][sm][lo:]; l = ohlcv['low'][sm][lo:]
            c = ohlcv['close'][sm][lo:]; v = ohlcv['vol'][sm][lo:]; amt = ohlcv['amount'][sm][lo:]
            sd2 = sd[sm][lo:]
            if len(c) < 60:
                continue
            # 对齐三序列收益(个股/行业/大盘); 对齐失败则 beta 留空, 技术因子仍算
            sret = _ret(c)
            mret = iret = np.array([])
            idx = stock_to_idx.get(code); ig = ind_groups.get(idx)
            if ig is not None:
                im = ig[0] <= d_str; in_ = int(im.sum()); ilo = max(0, in_ - maxlen)
                idf = pd.DataFrame({'trade_date': ig[0][im][ilo:], 'close': ig[1][im][ilo:]})
                mm = mkt_dates <= d_str; mn = int(mm.sum()); mlo = max(0, mn - maxlen)
                mdf = pd.DataFrame({'trade_date': mkt_dates[mm][mlo:], 'close': mkt_close[mm][mlo:]})
                if len(idf) >= 65 and len(mdf) >= 65:
                    sdf = pd.DataFrame({'trade_date': sd2, 'close': c})
                    try:
                        sr, kr, mr = _align_three(sdf, idf, mdf, maxlen)
                        if len(sr) >= 60:
                            sret, iret, mret = sr, kr, mr
                    except Exception:
                        pass
            # daily_basic 换手率/量比(≤报价日 切片, 喂 turnover_factors + 周月线)
            td, db = _get_daily_basic(code)
            turnover = vol_ratio_d = turnover_dates = None
            if td is not None:
                dbm = td <= d_str
                if dbm.any():
                    turnover = db['turnover'][dbm]
                    vol_ratio_d = db['vol_ratio'][dbm]
                    turnover_dates = td[dbm]
            f = compute_factors(o, h, l, c, v, amt, sret, mret, iret, dates=sd2,
                                turnover=turnover, vol_ratio=vol_ratio_d,
                                turnover_dates=turnover_dates)
            for k, val in f.items():
                all_factors.setdefault(k, np.full(len(df), np.nan))[i] = val
            fetched += 1
        except Exception:
            continue
        if (i + 1) % 100 == 0:
            print(f'    进度 {i+1}/{len(df)} | 有因子 {fetched}')

    for k, arr in all_factors.items():
        df[k] = arr
    cov = {k: f'{pd.Series(v).notna().mean()*100:.1f}%' for k, v in all_factors.items()}
    shown = ', '.join(f'{k}({cov[k]})' for k in sorted(all_factors)[:10])
    print(f'    匹配 {fetched}/{len(df)} | +{len(all_factors)}因子: {shown} ...')
    return df


# ====== J类: 月线10月均线趋势 (需 tushare 网络) ======

def derive_monthly_trend(df):
    """按每行报价日 PIT 回算 月线MA10_slope3% + 月线趋势向上, 写 2 列。

    bulk-load: 每 unique 股票 pro_bar(freq='M', adj='hfq') 全量1次(缓存), 切片 ≤报价日。
    PIT 由 ≤报价日 内存切片保证(只用已收盘完整月); hfq 保历史值不漂移。
    单股失败填 NaN/0, 不影响整体。
    """
    print('\n  J类: 月线10月均线趋势 (需tushare网络)...')
    pkg = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if pkg not in sys.path:
        sys.path.insert(0, pkg)
    from strategies.monthly_trend import trend10m
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())

    slope = np.full(len(df), np.nan)
    up = np.zeros(len(df), dtype=int)

    unique_codes = df['股票代码'].astype(str).unique().tolist()
    print(f'    取个股 hfq 月线全量: {len(unique_codes)} 只(缓存)...')
    cache = {}

    def get_monthly(code):
        if code in cache:
            return cache[code]
        try:
            d = ts.pro_bar(ts_code=code, freq='M', adj='hfq')
            if d is None or len(d) == 0:
                cache[code] = (None, None)
                return (None, None)
            d = d.sort_values('trade_date').drop_duplicates('trade_date').reset_index(drop=True)
            res = (d['trade_date'].astype(str).to_numpy(), d['close'].astype(float).to_numpy())
            cache[code] = res
            return res
        except Exception:
            cache[code] = (None, None)
            return (None, None)

    fetched = 0
    for i, row in enumerate(df.itertuples(index=True)):
        code = str(getattr(row, '股票代码', ''))
        d_raw = getattr(row, '报价日', None)
        if d_raw is None or pd.isna(d_raw):
            continue
        d_str = str(int(float(d_raw))) if isinstance(d_raw, (int, float, np.floating)) else str(d_raw)
        if len(d_str) < 8:
            continue
        try:
            dates, close = get_monthly(code)
            if dates is None:
                continue
            mask = dates <= d_str              # PIT: 只取 ≤报价日 的已收盘完整月
            if not mask.any():
                continue
            seg = close[mask][-40:]            # 末段40月, 足够算 MA10+滞后带状态
            if len(seg) < 13:
                continue
            r = trend10m(seg)
            if not np.isnan(r['slope3_pct']):
                slope[i] = r['slope3_pct']
                up[i] = r['trend_up']
            fetched += 1
        except Exception:
            continue
        if (i + 1) % 100 == 0:
            print(f'    进度 {i+1}/{len(df)} | 有信号 {fetched}')

    df['月线MA10_slope3%'] = pd.Series(slope, index=df.index)
    df['月线趋势向上'] = pd.Series(up, index=df.index)
    cov = f"{pd.notna(slope).mean()*100:.1f}%"
    up_rate = f"{up[pd.notna(slope)].mean()*100:.1f}%" if pd.notna(slope).any() else "n/a"
    print(f'    匹配 {fetched}/{len(df)} | +2特征: 月线MA10_slope3%(覆盖{cov}), 月线趋势向上(上行率{up_rate})')
    return df


# ====== K类: 大盘(沪深300)月线10月趋势 (需 tushare 网络, 全样本共享1次取数) ======

def derive_market_trend(df):
    """每样本按报价日 PIT 回算 大盘MA10_slope3%(沪深300月线10月均线3月斜率%), 写 1 列。

    标签是市场 regime 驱动 → 大盘自身10月趋势对准 regime, IV 0.18 且 train≈test 稳定
    (个股 backward-looking 则无信号, 见 J类)。全样本共享 1 次 index_daily 取数, 切片 ≤报价日。
    只产连续值(0/1 版 IV=0 不入模)。
    """
    print('\n  K类: 大盘月线10月趋势 (需tushare网络)...')
    pkg = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if pkg not in sys.path:
        sys.path.insert(0, pkg)
    from strategies.monthly_trend import trend10m
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    pro = ts.pro_api()

    # 沪深300 月线(指数无需复权): index_daily → 每月最后交易日收盘
    d = pro.index_daily(ts_code='000300.SH')
    d = d.sort_values('trade_date').copy()
    d['trade_date'] = d['trade_date'].astype(str)
    d['ym'] = d['trade_date'].str[:6]
    mon = d.groupby('ym').last().reset_index().sort_values('trade_date')
    md = mon['trade_date'].to_numpy()
    mc = mon['close'].astype(float).to_numpy()
    print(f'    沪深300 月线: {len(mon)} 月')

    slope = np.full(len(df), np.nan)
    fetched = 0
    for i, row in enumerate(df.itertuples(index=True)):
        d0 = getattr(row, '报价日', None)
        if d0 is None or pd.isna(d0):
            continue
        d_str = str(int(float(d0))) if isinstance(d0, (int, float, np.floating)) else str(d0)
        if len(d_str) < 8:
            continue
        mask = md <= d_str                # PIT: ≤报价日 的已收盘完整月
        if not mask.any():
            continue
        seg = mc[mask][-40:]
        if len(seg) < 13:
            continue
        r = trend10m(seg)
        if not np.isnan(r['slope3_pct']):
            slope[i] = r['slope3_pct']
            fetched += 1

    df['大盘MA10_slope3%'] = pd.Series(slope, index=df.index)
    print(f'    匹配 {fetched}/{len(df)} | +1特征: 大盘MA10_slope3%(覆盖{pd.notna(slope).mean()*100:.1f}%)')
    return df


# ====== 主流程 ======

def main():
    parser = argparse.ArgumentParser(description='定增特征衍生 - 从基础特征派生高级特征')
    parser.add_argument('input_path', nargs='?',
                        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'features.parquet'),
                        help='输入特征文件 (parquet或CSV)')
    parser.add_argument('--output', default=None,
                        help='输出路径 (默认: 同目录下 features_derived.parquet)')
    parser.add_argument('--no-db', action='store_true', help='跳过需MySQL的特征(F+G类)')
    args = parser.parse_args()

    # 输出路径
    if args.output:
        output_path = args.output
    else:
        base_dir = os.path.dirname(args.input_path)
        output_path = os.path.join(base_dir, 'features_derived.parquet')

    # 加载
    print(f'加载基础特征: {args.input_path}')
    if args.input_path.endswith('.parquet'):
        df = pd.read_parquet(args.input_path)
    else:
        df = pd.read_csv(args.input_path)
    orig_cols = len(df.columns)
    print(f'  {len(df)} 行 × {orig_cols} 列')

    # ====== Stage 1: Parquet-only ======
    print('\n' + '='*60)
    print('Stage 1: Parquet-only 衍生特征')
    print('='*60)

    df = derive_fcf_growth_rates(df)
    df = derive_fcf_cross_metrics(df)
    df = derive_financial_score_deltas(df)
    df = derive_valuation_relative(df)
    df = derive_market_momentum(df)
    df = derive_placement_structure(df)

    # ====== Stage 2: DB-dependent ======
    if not args.no_db:
        print('\n' + '='*60)
        print('Stage 2: DB-dependent 衍生特征')
        print('='*60)
        try:
            df = derive_industry_valuation_growth(df)
            df = derive_market_index_features(df)
        except Exception as e:
            print(f'  ⚠️ DB特征失败: {e}')

        # D类补: PB_vs_同行中位 ← PIT 行业口径(覆盖 Stage1 的 peer 中位; 与回测同源; 需 daily_basic 网络)
        try:
            df = derive_pb_vs_industry_pit(df)
        except Exception as e:
            print(f'  ⚠️ PB PIT 行业口径失败: {e}(保留原 peer 口径)')

        # H/I类: 三浪/抵抗策略信号(需tushare网络, 单独try, 失败不影响F/G)
        try:
            df = derive_strategy_signals(df)
        except Exception as e:
            import traceback
            print(f'  ⚠️ 策略信号(H/I类)失败: {e}')
            traceback.print_exc()

        # L类: Beta+Alpha158 因子(需tushare网络, 复用 H/I 已热的 OHLCV 缓存, 失败不影响前面)
        try:
            df = derive_alpha_beta_factors(df)
        except Exception as e:
            import traceback
            print(f'  ⚠️ 因子引擎(L类)失败: {e}')
            traceback.print_exc()

        # J类: 月线10月均线趋势(需tushare网络, 单独try, 失败不影响前面)
        try:
            df = derive_monthly_trend(df)
        except Exception as e:
            import traceback
            print(f'  ⚠️ 月线趋势(J类)失败: {e}')
            traceback.print_exc()

        # K类: 大盘(沪深300)月线10月趋势(需tushare网络, 单独try; 标签regime驱动→有信号, 入模)
        try:
            df = derive_market_trend(df)
        except Exception as e:
            import traceback
            print(f'  ⚠️ 大盘趋势(K类)失败: {e}')
            traceback.print_exc()

    # ====== 清理 ======
    df = df.replace([np.inf, -np.inf], np.nan)

    # ====== 末尾闸门：剔除确认无信号的字段（IV=0 标记/常数/标识）======
    # 放在最后一步，拦住本脚本自己重新生成的字段（如 市场_above_MA250）
    DROP_FIELDS = {
        # 子场景通过标记（近常数，IV=0）
        '市场指数_通过', '行业PE_通过', '个股PE_通过', 'DCF估值_通过',
        '修正PE估值_通过', '参数构造_通过', '蒙特卡洛_通过', '反向推算_通过',
        '子场景通过数', 'step1通过', 'step2通过', 'step3通过',
        # 日期/行情标识（IV=0，不该入模）
        '最新交易日', '行情_时间匹配', '财报年份', 'report_year',
        # derive 重新生成的无信号字段
        '市场_above_MA250',
        # 定增参数（IV=0）
        '溢价率', '溢价率下限', '锁定期', '融资金额', '无风险利率', 'Beta',
        '定增建议参与', '有效阈值数',
        # 行业代码文本（不该入模）
        'sw_l1_code', 'sw_l1_name', 'sw_l2_code', 'sw_l2_name', 'sw_l3_code', 'sw_l3_name',
        # 稀疏 FCF 营运资金变动
        '营运资金变动_T', '营运资金变动_T1', '营运资金变动_T2', '营运资金变动_T3', '营运资金变动_T4',
        # 弱/重复
        '营业利润率', '营收增长率', '行业PS', '净债务', '净资产负债表',
        # 行业级别：1391个唯一值(几乎每行不同)，非干净分级，脏数据
        '行业级别',
    }
    drop_present = [c for c in DROP_FIELDS if c in df.columns]
    if drop_present:
        df = df.drop(columns=drop_present)
        print(f'  末尾剔除无信号字段: {len(drop_present)} 个')

    # ====== 保存 ======
    new_cols_count = len(df.columns) - orig_cols
    df.to_parquet(output_path, index=False)
    csv_path = output_path.replace('.parquet', '.csv')
    df.to_csv(csv_path, index=False)

    print('\n' + '='*60)
    print(f'✅ 衍生特征已保存: {output_path}')
    print(f'   {len(df)} 行 × {len(df.columns)} 列 (原始{orig_cols} + 新增{new_cols_count})')
    print(f'   CSV(查看用): {csv_path}')

    # 新增特征覆盖率汇总
    new_col_names = df.columns[orig_cols:].tolist()
    if new_col_names:
        print(f'\n新增特征覆盖率:')
        for c in new_col_names:
            rate = df[c].notna().mean() * 100
            print(f'  {c:30s} {rate:5.1f}%')

    # ── 冻结快照入 DB 版本库(训练真正吃的就是这个 derived 矩阵) ──
    try:
        import db_dataset_store
        derived_version = db_dataset_store.save_snapshot(
            df, kind='derived', label_config='7m',
            note=f'derived {len(df)}x{len(df.columns)} (orig{orig_cols}+new{new_cols_count})')
        print(f'   快照入DB: {derived_version}')
    except Exception as e:
        print(f'   ⚠️ 快照入库跳过(DB不可用): {e}')


if __name__ == '__main__':
    main()
