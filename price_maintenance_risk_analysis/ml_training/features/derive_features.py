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

    # 评分斜率(polyfit T..T-4 vs year; 恢复历史 总分_斜率/盈利能力_斜率/成长能力_斜率, 1w 模型用)
    _years = np.array([-4, -3, -2, -1, 0]); _yc = _years - _years.mean()
    _suf = ['_T-4', '_T-3', '_T-2', '_T-1', '_T']
    for m in SCORE_METRICS:
        cols_m = [f'{m}{s}' for s in _suf]
        if not all(c in df.columns for c in cols_m):
            continue
        vals = np.column_stack([pd.to_numeric(df[c], errors='coerce').to_numpy() for c in cols_m])  # (n,5)
        slope = np.full(len(df), np.nan)
        full = ~np.isnan(vals).any(axis=1)
        if full.any():
            slope[full] = (vals[full] * _yc).sum(axis=1) / (_yc @ _yc)   # 向量化: slope=Σ(yc·y)/Σ(yc²)
        partial = (~full) & (~np.isnan(vals).all(axis=1))
        for i in np.where(partial)[0]:
            row = vals[i]; mask = ~np.isnan(row)
            if mask.sum() >= 2:
                slope[i] = np.polyfit(_years[mask], row[mask], 1)[0]
        new_cols[f'{m}_斜率'] = pd.Series(slope, index=df.index)

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in list(coverage.items())[:6])}...')
    return df



# ====== D类预: 个股估值 PIT + 同行截面均值(恢复 load_db_features/peer_companies 丢失的 个股PE/PS + 同行) ======

_STOCK_TO_IDX = None   # lazy: stock_code → index_code(行业映射, industry_data 一次查)


def _load_stock_to_idx():
    """stock_code → 申万行业 index_code(industry_data)。模块级缓存, 首次查 DB。"""
    global _STOCK_TO_IDX
    if _STOCK_TO_IDX is not None:
        return _STOCK_TO_IDX
    try:
        import pymysql
        from utils.db_manager import ValuationDB
        conn = pymysql.connect(**ValuationDB.MYSQL_CONFIG)
        idf = pd.read_sql('SELECT stock_code, index_code FROM industry_data WHERE index_code IS NOT NULL', conn)
        conn.close()
        _STOCK_TO_IDX = dict(zip(idf['stock_code'].astype(str), idf['index_code'].astype(str)))
    except Exception:
        _STOCK_TO_IDX = {}
    return _STOCK_TO_IDX


def derive_peer_valuation(df):
    """个股 PE/PS/PB/市值(PIT, daily_basic ≤报价日) + 同行均值/中位(截面 groupby 报价日×行业)。
    恢复历史 load_db_features(relative_valuation+peer_companies 快照, 非PIT)丢失的估值特征;
    本函数改从 daily_basic 时序表 PIT 切(≤报价日), 同行=截面均值(全A 含全成员)。
    喂 derive_valuation_relative 产 PE_vs_行业/同行、PB_vs、PS_vs。需先 prefetch_daily_basic。"""
    print('\n  D类预: 个股估值PIT + 同行截面均值...')
    if '股票代码' not in df.columns or '报价日' not in df.columns:
        return df
    s2i = _load_stock_to_idx()
    cols_map = {'个股PE': 'pe', '个股PS': 'ps', '个股PB': 'pb', '个股市值': 'total_mv'}
    out = {c: np.full(len(df), np.nan) for c in cols_map}
    dates = df['报价日'].astype(str).to_numpy()
    n_hit = 0
    for code, g in df.groupby('股票代码'):
        cached = _DAILY_BASIC_CACHE.get(str(code))
        if cached is None or cached[0] is None:
            continue
        sd = cached[0]; dbd = cached[1]
        for idx in g.index:
            pos = int(np.searchsorted(sd, dates[idx], 'right')) - 1
            if pos < 0:
                continue
            n_hit += 1
            for cn, dn in cols_map.items():
                arr = dbd.get(dn)
                if arr is not None and pos < len(arr):
                    out[cn][idx] = arr[pos]
    for cn in cols_map:
        df[cn] = out[cn]
    # 同行截面均值/中位(groupby 报价日×行业 → transform; 全A 含全成员)
    df['_行业idx'] = df['股票代码'].astype(str).map(s2i)
    valid = df['_行业idx'].notna()
    for cn in ['个股PE', '个股PS', '个股PB', '个股市值']:
        short = cn[2:]   # PE/PS/PB/市值
        df['同行' + short + '_均值'] = np.nan
        df['同行' + short + '_中位'] = np.nan
        if valid.any():
            df.loc[valid, '同行' + short + '_均值'] = df[valid].groupby(['报价日', '_行业idx'])[cn].transform('mean')
            df.loc[valid, '同行' + short + '_中位'] = df[valid].groupby(['报价日', '_行业idx'])[cn].transform('median')
    df = df.drop(columns=['_行业idx'])
    cov = {c: f'{df[c].notna().mean()*100:.0f}%' for c in ['个股PE', '个股PS', '同行PS_均值', '同行市值_均值']}
    print(f'    +12特征(个股估值4+同行均值4+中位4, n_hit{n_hit}): {", ".join(f"{k}({v})" for k,v in cov.items())}')
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
    新口径 = daily_basic PB(≤报价日) / industry_daily sw_index_pb(≤报价日), 与回测 export_features(PIT loader)同源。
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
    from features.export_features import load_pb_vs_industry

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


# ====== E类预: 行情统计(从 _OHLCV_CACHE PIT 算) → 喂 derive_market_momentum ======

def derive_market_stats_from_ohlcv(df):
    """从 _OHLCV_CACHE PIT 算行情统计基列: MA20/30/60/120/250、波动率/年化收益/区间收益/胜率_20-250d、
    当前价、漂移率、波动率(全期)、中位价、价格标准差、数据天数(共 ~27 个)。
    每股全历史 rolling/expanding 一次, 按报价日 searchsorted(≤报价日)索引。PIT 正确(≤D 序列)。

    恢复: 历史 load_db_features 读 market_data 表(每股一行非PIT快照, 有泄漏)产这些列; 回测没接 → 丢失。
    本函数从已预取的 _OHLCV_CACHE 重算(PIT), 喂给 derive_market_momentum 产 vol_ratio/return_acceleration/price_vs_MA。
    缺缓存(如定增侧未 prefetch_ohlcv)→ 静默跳过(定增走 market_data 表另得)。"""
    print('\n  E类预: 行情统计(从 OHLCV PIT)...')
    if '股票代码' not in df.columns or '报价日' not in df.columns:
        print('    缺 股票代码/报价日, 跳过'); return df
    windows = [20, 30, 60, 120, 250]
    cols = ([f'MA{w}' for w in windows] + [f'波动率_{w}d' for w in windows] +
            [f'年化收益_{w}d' for w in windows] + [f'区间收益_{w}d' for w in windows] +
            [f'胜率_{w}d' for w in windows] +
            ['当前价', '漂移率', '波动率', '中位价', '价格标准差', '数据天数'])
    out = {c: np.full(len(df), np.nan) for c in cols}
    dates = df['报价日'].astype(str).to_numpy()
    n_hit = 0
    for code, g in df.groupby('股票代码'):
        cached = _OHLCV_CACHE.get(str(code))
        if cached is None or cached[0] is None:
            continue
        sd = cached[0]; close = cached[1]['close'].astype(float)
        n = len(sd)
        if n < 2:
            continue
        cs = pd.Series(close)
        rs = cs.pct_change()                      # rets 对齐 close 轴, rets[0]=nan
        stat = {}
        for w in windows:
            stat[f'MA{w}'] = cs.rolling(w).mean().to_numpy()
            stat[f'波动率_{w}d'] = rs.rolling(w).std(ddof=0).to_numpy() * np.sqrt(250)
            prev = cs.shift(w).to_numpy()
            ratio = cs.to_numpy() / np.where(prev > 0, prev, np.nan)
            stat[f'年化收益_{w}d'] = np.where(prev > 0, ratio ** (250.0 / w) - 1, np.nan)
            stat[f'区间收益_{w}d'] = np.where(prev > 0, ratio - 1, np.nan)
            stat[f'胜率_{w}d'] = pd.Series((rs > 0).to_numpy().astype(float)).rolling(w).mean().to_numpy()
        stat['当前价'] = close
        stat['漂移率'] = rs.expanding().mean().to_numpy() * 250
        stat['波动率'] = rs.expanding().std(ddof=0).to_numpy() * np.sqrt(250)
        stat['中位价'] = cs.expanding().median().to_numpy()
        stat['价格标准差'] = cs.expanding().std(ddof=0).to_numpy()
        stat['数据天数'] = np.arange(1, n + 1, dtype=float)
        for idx in g.index:
            pos = int(np.searchsorted(sd, dates[idx], 'right')) - 1
            if pos < 0:
                continue
            n_hit += 1
            for c, arr in stat.items():
                out[c][idx] = arr[pos]
    for c in cols:
        df[c] = out[c]
    cov = {c: f'{pd.Series(out[c]).notna().mean()*100:.0f}%' for c in cols[:8]}
    print(f'    +{len(cols)}个特征 (n_hit行 {n_hit}/{len(df)}): {", ".join(f"{k}({v})" for k,v in cov.items())}...')
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
            'close': g['close'].values.astype(float),   # 行业指数收盘(算 行业年化收益/区间收益/波动率/胜率)
        }

    # 获取每条样本的 index_code
    sample_index_codes = df['股票代码'].map(stock_to_index)

    # 对每条样本计算
    growth_windows = [60, 120, 250]
    pe_growth = {w: np.full(len(df), np.nan) for w in growth_windows}
    pb_growth = {w: np.full(len(df), np.nan) for w in growth_windows}
    # 行业指数收益统计(从 close, PIT ≤报价日; 恢复历史 load_db_features 丢失的 行业年化收益/区间收益)
    ind_windows = [20, 60, 120, 250]
    ind_range = {w: np.full(len(df), np.nan) for w in ind_windows}   # 行业区间收益_{w}d
    ind_ann = {w: np.full(len(df), np.nan) for w in ind_windows}     # 行业年化收益_{w}d
    ind_pe = np.full(len(df), np.nan)   # 行业PE 水平(≤报价日)
    ind_pb = np.full(len(df), np.nan)   # 行业PB 水平(≤报价日)

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
        ind_pe[i] = pe_now   # 行业PE/PB 水平(≤报价日; 恢复历史 load_db_features 的 行业PE/行业PB)
        ind_pb[i] = pb_now
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
        # 行业指数收益(从 close, PIT ≤报价日; 恢复历史 load_db_features 丢失的 行业年化/区间收益)
        ind_close = dg['close']
        for w in ind_windows:
            if pos >= w:
                c0 = ind_close[pos - w]; c1 = ind_close[pos]
                if not np.isnan(c0) and not np.isnan(c1) and abs(c0) > 1e-8:
                    ind_range[w][i] = c1 / c0 - 1
                    ind_ann[w][i] = (c1 / c0) ** (250.0 / w) - 1
        matched += 1

    print(f'    匹配行业日线: {matched}/{len(df)}')

    new_cols = {}
    for w in growth_windows:
        new_cols[f'行业PE_{w}d增长'] = pd.Series(pe_growth[w], index=df.index)
        new_cols[f'行业PB_{w}d增长'] = pd.Series(pb_growth[w], index=df.index)
    for w in ind_windows:
        new_cols[f'行业区间收益_{w}d'] = pd.Series(ind_range[w], index=df.index)
        new_cols[f'行业年化收益_{w}d'] = pd.Series(ind_ann[w], index=df.index)
    new_cols['行业PE'] = pd.Series(ind_pe, index=df.index)
    new_cols['行业PB'] = pd.Series(ind_pb, index=df.index)

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


def _f(x):
    """float, NaN/inf → None。pymysql executemany 遇 nan 抛 ProgrammingError 整批失败(被各 save 的 except 吞), 故落盘前必清洗。MySQL DOUBLE 接受 NULL。"""
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    if v != v or v in (float('inf'), float('-inf')):   # nan / inf
        return None
    return v


def _bulk_group_split(df):
    """df(含 stock_code + trade_date) → 全局排序后按 stock_code 用 numpy 切边界。
    替代逐组 groupby+sort_values+逐组 .values: 11.6M 行 object dtype 上 take_2d_axis1 爆慢(全A曾卡 12min+)。
    返回 (排序后 df, 各组行索引 ndarray list)。调用方按 idx 切片取列(numpy fancy index, 快)。"""
    df = df.sort_values(['stock_code', 'trade_date'], kind='mergesort')
    sc = df['stock_code'].to_numpy()
    n = len(sc)
    change = np.flatnonzero(sc[1:] != sc[:-1]) + 1 if n > 1 else np.array([], dtype=int)
    return df, np.split(np.arange(n), change)


def _qfq_bulk_read(codes, chunk=800):
    """从 stock_qfq_daily 批量读 {code: (dates, ohlcv)}。表不存在/空→{}。首次重建后零 tushare。
    分块读: 全A 11.6M 行一次性 read_sql 峰值 ~3GB 致 8GB Mac swap 卡死; 每块 ~800 股, 峰值内存有界。"""
    if not codes:
        return {}
    out = {}
    try:
        import pymysql
        from utils.db_manager import ValuationDB
        cfg = ValuationDB.MYSQL_CONFIG
        conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                               password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
        try:
            for i in range(0, len(codes), chunk):
                blk = codes[i:i + chunk]
                ph = ','.join(['%s'] * len(blk))
                df = pd.read_sql(f"SELECT stock_code,trade_date,open,high,low,close,vol,amount "
                                 f"FROM stock_qfq_daily WHERE stock_code IN ({ph})", conn, params=list(blk))
                if df is None or df.empty:
                    continue
                df, splits = _bulk_group_split(df)
                sc = df['stock_code'].to_numpy(); dates = df['trade_date'].to_numpy()
                o = df['open'].to_numpy(); h = df['high'].to_numpy(); l = df['low'].to_numpy()
                c = df['close'].to_numpy(); v = df['vol'].to_numpy(); amt = df['amount'].to_numpy()
                for idx in splits:
                    if not len(idx):
                        continue
                    out[str(sc[idx[0]])] = (dates[idx], {'open': o[idx], 'high': h[idx], 'low': l[idx],
                        'close': c[idx], 'vol': v[idx], 'amount': amt[idx]})
        finally:
            conn.close()
    except Exception:
        pass
    return out


def _qfq_save(code, dates, ohlcv):
    """一只股的 qfq 全历史落盘 stock_qfq_daily(ON DUP KEY, 建表幂等)。落盘失败不影响计算。"""
    if dates is None or len(dates) == 0:
        return
    try:
        import pymysql
        from utils.db_manager import ValuationDB
        cfg = ValuationDB.MYSQL_CONFIG
        conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                               password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
        n = len(dates)
        rows = [(code, str(dates[i]), _f(ohlcv['open'][i]), _f(ohlcv['high'][i]), _f(ohlcv['low'][i]),
                 _f(ohlcv['close'][i]), _f(ohlcv['vol'][i]), _f(ohlcv['amount'][i])) for i in range(n)]
        try:
            with conn.cursor() as cur:
                cur.execute("""CREATE TABLE IF NOT EXISTS stock_qfq_daily (
                    stock_code VARCHAR(16) NOT NULL, trade_date CHAR(8) NOT NULL,
                    open DOUBLE, high DOUBLE, low DOUBLE, close DOUBLE, vol DOUBLE, amount DOUBLE,
                    PRIMARY KEY (stock_code, trade_date)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
                sql = ("INSERT INTO stock_qfq_daily (stock_code,trade_date,open,high,low,close,vol,amount) "
                       "VALUES (%s,%s,%s,%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE "
                       "open=VALUES(open),high=VALUES(high),low=VALUES(low),close=VALUES(close),vol=VALUES(vol),amount=VALUES(amount)")
                B = 2000
                for i in range(0, n, B):
                    cur.executemany(sql, rows[i:i + B])
            conn.commit()
        finally:
            conn.close()
    except Exception:
        pass


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
    优先级: 内存缓存 → stock_qfq_daily 本地表 → pro_bar+落盘。
    供 derive_strategy_signals(取 close) 与 derive_alpha_beta_factors(OHLCV) 共用。"""
    if code in _OHLCV_CACHE:
        return _OHLCV_CACHE[code]
    local = _qfq_bulk_read([code])          # 本地表(免 pro_bar)
    if code in local:
        _OHLCV_CACHE[code] = local[code]
        return local[code]
    try:
        res = _fetch_ohlcv_raw(code)
        _qfq_save(code, res[0], res[1])     # 落盘供下次
    except Exception:
        res = (None, None)
    _OHLCV_CACHE[code] = res
    return res


def prefetch_ohlcv(codes, max_workers=12):
    """并发预取所有 unique 股票 OHLCV 入缓存(I/O bound, 线程即可)。
    本地优先(stock_qfq_daily 批量读), 缺失的才 pro_bar+落盘。tushare 限流单股退避重试(3 次)。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import time
    uniq = [c for c in dict.fromkeys(str(c) for c in codes) if c not in _OHLCV_CACHE]
    if not uniq:
        print(f'    OHLCV 缓存已热({len(_OHLCV_CACHE)}), 跳过预取')
        return
    # 1. 本地 stock_qfq_daily 批量读(免 tushare; 首次重建后命中全部)
    local = _qfq_bulk_read(uniq)
    for c, res in local.items():
        _OHLCV_CACHE[c] = res
    todo = [c for c in uniq if _OHLCV_CACHE.get(c, (None, None))[0] is None]
    if not todo:
        print(f'    OHLCV 全本地命中 {len(local)} 只, 零 tushare')
        return
    print(f'    OHLCV: 本地 {len(local)}, 待 pro_bar {len(todo)} (max_workers={max_workers}, 落盘)...')

    def _work(code):
        for attempt in range(3):
            try:
                res = _fetch_ohlcv_raw(code)
                _OHLCV_CACHE[code] = res
                _qfq_save(code, res[0], res[1])      # 落盘 stock_qfq_daily 供下次重建
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
    print(f'    预取完成: {ok}/{len(todo)} 成功(已落盘 stock_qfq_daily)')


# ====== 模块级 daily_basic 缓存(换手率/量比, 与 OHLCV 缓存并列) ======
_DAILY_BASIC_CACHE = {}   # code -> (dates, {turnover, vol_ratio}) 或 (None, None)


def _db_basic_bulk_read(codes):
    """从 stock_daily_basic 批量读 {code: (dates, {turnover, vol_ratio})}。表空→{}。"""
    if not codes:
        return {}
    try:
        import pymysql
        from utils.db_manager import ValuationDB
        cfg = ValuationDB.MYSQL_CONFIG
        conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                               password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
        try:
            ph = ','.join(['%s'] * len(codes))
            df = pd.read_sql(f"SELECT stock_code,trade_date,turnover_rate,volume_ratio,pb,pe,ps,total_mv "
                             f"FROM stock_daily_basic WHERE stock_code IN ({ph})", conn, params=list(codes))
        finally:
            conn.close()
    except Exception:
        return {}
    if df is None or df.empty:
        return {}
    out = {}
    df, splits = _bulk_group_split(df)
    sc = df['stock_code'].to_numpy(); dates = df['trade_date'].to_numpy()
    def _arr(nm): return df[nm].to_numpy() if nm in df.columns else None
    to = _arr('turnover_rate'); vr = _arr('volume_ratio'); pb = _arr('pb')
    pe = _arr('pe'); ps = _arr('ps'); mv = _arr('total_mv')
    for idx in splits:
        if not len(idx):
            continue
        out[str(sc[idx[0]])] = (dates[idx], {'turnover': to[idx], 'vol_ratio': vr[idx],
            'pb': pb[idx] if pb is not None else None,
            'pe': pe[idx] if pe is not None else None,
            'ps': ps[idx] if ps is not None else None,
            'total_mv': mv[idx] if mv is not None else None})
    return out


def _db_basic_save(code, dates, db):
    """一只股 daily_basic 全历史落盘 stock_daily_basic(ON DUP KEY, 建表幂等)。"""
    if dates is None or len(dates) == 0:
        return
    try:
        import pymysql
        from utils.db_manager import ValuationDB
        cfg = ValuationDB.MYSQL_CONFIG
        conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                               password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
        n = len(dates)
        rows = [(code, str(dates[i]), _f(db['turnover'][i]), _f(db['vol_ratio'][i]),
                 _f(db['pb'][i]) if db.get('pb') is not None else None,
                 _f(db['pe'][i]) if db.get('pe') is not None else None,
                 _f(db['ps'][i]) if db.get('ps') is not None else None,
                 _f(db['total_mv'][i]) if db.get('total_mv') is not None else None) for i in range(n)]
        try:
            with conn.cursor() as cur:
                cur.execute("""CREATE TABLE IF NOT EXISTS stock_daily_basic (
                    stock_code VARCHAR(16) NOT NULL, trade_date CHAR(8) NOT NULL,
                    turnover_rate DOUBLE, volume_ratio DOUBLE, pb DOUBLE,
                    pe DOUBLE, ps DOUBLE, total_mv DOUBLE,
                    PRIMARY KEY (stock_code, trade_date)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
                # 旧表补列(pb/pe/ps/total_mv; 幂等)
                for col in ('pb', 'pe', 'ps', 'total_mv'):
                    cur.execute("""SELECT COUNT(*) FROM information_schema.COLUMNS
                        WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='stock_daily_basic' AND COLUMN_NAME=%s""", (col,))
                    if cur.fetchone()[0] == 0:
                        cur.execute(f"ALTER TABLE stock_daily_basic ADD COLUMN {col} DOUBLE")
                sql = ("INSERT INTO stock_daily_basic (stock_code,trade_date,turnover_rate,volume_ratio,pb,pe,ps,total_mv) "
                       "VALUES (%s,%s,%s,%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE "
                       "turnover_rate=VALUES(turnover_rate),volume_ratio=VALUES(volume_ratio),pb=VALUES(pb),"
                       "pe=VALUES(pe),ps=VALUES(ps),total_mv=VALUES(total_mv)")
                B = 2000
                for i in range(0, n, B):
                    cur.executemany(sql, rows[i:i + B])
            conn.commit()
        finally:
            conn.close()
    except Exception:
        pass


def _fetch_daily_basic_raw(code):
    """实际调 daily_basic(无缓存), 返回 (dates, {turnover, vol_ratio}); 失败抛异常。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    d = ts.pro_api().daily_basic(ts_code=code,
                                 fields='ts_code,trade_date,turnover_rate,volume_ratio,pb,pe,ps,total_mv')
    if d is None or len(d) == 0:
        raise ValueError(f'daily_basic 空返回: {code}')
    d = d.sort_values('trade_date').drop_duplicates('trade_date').reset_index(drop=True)
    def _col(nm):
        return d[nm].astype(float).values if nm in d.columns else np.full(len(d), np.nan)
    return (d['trade_date'].values,
            {'turnover': _col('turnover_rate'), 'vol_ratio': _col('volume_ratio'),
             'pb': _col('pb'), 'pe': _col('pe'), 'ps': _col('ps'), 'total_mv': _col('total_mv')})


def _get_daily_basic(code):
    """daily_basic 全量(缓存), 返回 (dates, {turnover, vol_ratio}) 或 (None, None)。
    优先级: 内存缓存 → stock_daily_basic 本地表 → tushare+落盘。"""
    if code in _DAILY_BASIC_CACHE:
        return _DAILY_BASIC_CACHE[code]
    local = _db_basic_bulk_read([code])           # 本地表(免 tushare)
    if code in local:
        _DAILY_BASIC_CACHE[code] = local[code]
        return local[code]
    try:
        res = _fetch_daily_basic_raw(code)
        _db_basic_save(code, res[0], res[1])      # 落盘供下次
    except Exception:
        res = (None, None)
    _DAILY_BASIC_CACHE[code] = res
    return res


def prefetch_daily_basic(codes, max_workers=12):
    """并发预取所有 unique 股票 daily_basic 入缓存。
    本地优先(stock_daily_basic 批量读), 缺失的才 tushare+落盘。限流单股退避重试(3 次)。"""
    from concurrent.futures import ThreadPoolExecutor, as_completed
    import time
    uniq = [c for c in dict.fromkeys(str(c) for c in codes) if c not in _DAILY_BASIC_CACHE]
    if not uniq:
        print(f'    daily_basic 缓存已热({len(_DAILY_BASIC_CACHE)}), 跳过预取')
        return
    local = _db_basic_bulk_read(uniq)
    for c, res in local.items():
        _DAILY_BASIC_CACHE[c] = res
    todo = [c for c in uniq if _DAILY_BASIC_CACHE.get(c, (None, None))[0] is None]
    if not todo:
        print(f'    daily_basic 全本地命中 {len(local)} 只, 零 tushare')
        return
    print(f'    daily_basic: 本地 {len(local)}, 待 tushare {len(todo)} (max_workers={max_workers}, 落盘)...')

    def _work(code):
        for attempt in range(3):
            try:
                res = _fetch_daily_basic_raw(code)
                _DAILY_BASIC_CACHE[code] = res
                _db_basic_save(code, res[0], res[1])   # 落盘 stock_daily_basic
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
    print(f'    预取完成: {ok}/{len(todo)} 成功(已落盘 stock_daily_basic)')


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
    _orig_idx = df.index
    df = df.reset_index(drop=True)   # all_factors 按位置索引; fork chunk 标签非连续(iloc[i::n])→ 重置 0..n-1=位置, 末尾恢复(免 IndexError)
    print('\n  L类: Beta+Alpha158 因子(factor_engine)...')
    pkg = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if pkg not in sys.path:
        sys.path.insert(0, pkg)
    from strategies.data_loader import _align_three, _MARKET_INDEX
    from features.factor_engine import compute_factors, _ret

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
    # ===== 按股全历史算一次 SERIES + 按报价日索引(替代逐行重切; 日线族+beta 精确向量化,
    #       multiperiod/smc 每股 resample 缓存一次 + 逐行建 ≤D 序列含当周部分 bar, PIT 精确复刻标量) =====
    from features.factor_engine import (compute_factors_series, _beta_series, _mp_indicators,
                               _resample_ranges, smc_factors, _FACTOR_NAMES_OF)
    from strategies.data_loader import _align_three_full

    all_factors = {}
    _all_names = (_FACTOR_NAMES_OF['kline_factors'] + _FACTOR_NAMES_OF['tech_factors']
                  + _FACTOR_NAMES_OF['volume_factors'] + _FACTOR_NAMES_OF['moment_factors']
                  + _FACTOR_NAMES_OF['beta_factors'] + _FACTOR_NAMES_OF['smc_factors']
                  + _FACTOR_NAMES_OF['turnover_factors'] + _FACTOR_NAMES_OF['multiperiod_factors']
                  + _FACTOR_NAMES_OF['smc_factors_multiperiod'] + _FACTOR_NAMES_OF['turnover_multiperiod'])
    for k in _all_names:
        all_factors[k] = np.full(len(df), np.nan)

    uniq = df['股票代码'].astype(str).unique()
    done = 0
    for code, grp in df.groupby('股票代码'):
        code = str(code)
        sd, ohlcv = _get_stock_ohlcv(code)
        if sd is None:
            continue
        sd = sd.astype(str)
        o = ohlcv['open']; h = ohlcv['high']; l = ohlcv['low']; c = ohlcv['close']; v = ohlcv['vol']; amt = ohlcv['amount']
        td, db = _get_daily_basic(code)
        turnover = db['turnover'] if td is not None else None
        vol_ratio = db['vol_ratio'] if td is not None else None
        daily_ser = compute_factors_series(o, h, l, c, v, amt)   # 日线族(不含 turnover, td 轴另算)
        from features.factor_engine import _turnover_series as _tser
        td_ser = _tser(turnover, vol_ratio) if td is not None else {}
        td_str = td.astype(str) if td is not None else None
        # turnover_multiperiod: td 轴 W/M 日均换手率(每股 resample 一次; 完成柱口径, 非模型特征)
        tmp = {}
        if td is not None and len(td) >= 15:
            tidx = pd.to_datetime(pd.Series(td_str), format='%Y%m%d')
            ts = pd.Series(np.asarray(turnover, float), index=tidx)
            for rule, tag, nbars in [('W', 'W', 13), ('M', 'M', 6)]:
                r = ts.resample(rule).mean().dropna()
                tmp[tag] = (r.index.strftime('%Y%m%d').to_numpy(), r.values, nbars)
        # beta: 全量对齐一次 → align 轴 series(按报价日 ≤D 取末 N)
        beta_ser = {}; adates = np.array([])
        idx = stock_to_idx.get(code); ig = ind_groups.get(idx)
        if ig is not None:
            try:
                sdf = pd.DataFrame({'trade_date': sd, 'close': c})
                idf = pd.DataFrame({'trade_date': ig[0], 'close': ig[1]})
                mdf = pd.DataFrame({'trade_date': mkt_dates, 'close': mkt_close})
                adates, sret, mret, iret = _align_three_full(sdf, idf, mdf)
                if len(adates):
                    beta_ser = _beta_series(sret, mret, iret)
            except Exception:
                pass
        # multiperiod/smc_mp resample 缓存(每股一次, 复用日线起止索引)
        Wr = _resample_ranges(sd, o, h, l, c, 'W') if len(sd) >= 15 else None
        Mr = _resample_ranges(sd, o, h, l, c, 'M') if len(sd) >= 15 else None

        for row_idx, d_str in zip(grp.index.values, grp['报价日'].astype(str).values):
            pos = int(np.searchsorted(sd, d_str, 'right')) - 1
            if pos < 0:
                continue
            for k, arr in daily_ser.items():                      # 日线因子(close 轴)
                if pos < len(arr) and not np.isnan(arr[pos]):
                    all_factors[k][row_idx] = arr[pos]
            if td_ser:                                            # turnover 族(td 轴, 按 td-pos)
                tpos = int(np.searchsorted(td_str, d_str, 'right')) - 1
                if tpos >= 0:
                    for k, arr in td_ser.items():
                        if tpos < len(arr) and not np.isnan(arr[tpos]):
                            all_factors[k][row_idx] = arr[tpos]
            if beta_ser:                                          # beta(align 轴)
                apos = int(np.searchsorted(adates, d_str, 'right')) - 1
                if apos >= 0:
                    for k, arr in beta_ser.items():
                        if apos < len(arr) and not np.isnan(arr[apos]):
                            all_factors[k][row_idx] = arr[apos]
            if pos >= 59:                                         # smc 日线(≤D last 800)
                lo = max(0, pos + 1 - maxlen)
                try:
                    smc = smc_factors(o[lo:pos + 1], h[lo:pos + 1], l[lo:pos + 1], c[lo:pos + 1])
                    for k, val in smc.items():
                        if not np.isnan(val):
                            all_factors[k][row_idx] = val
                except Exception:
                    pass
            for rng, tag, NM in [(Wr, 'W', (48, 12)), (Mr, 'M', (24, 6))]:   # multiperiod + smc_mp
                if rng is None:
                    continue
                last_date, fidx = rng['last_date'], rng['first_idx']
                bc = int(np.searchsorted(last_date, d_str, 'right')) - 1
                if bc < 0:
                    continue
                wo = rng['wo'][:bc + 1]; wh = rng['wh'][:bc + 1]; wl = rng['wl'][:bc + 1]; wc = rng['wc'][:bc + 1]
                partial = False
                if bc + 1 < len(rng['wc']) and sd[fidx[bc + 1]] <= d_str:
                    s = fidx[bc + 1]; ep = pos
                    seg_h = h[s:ep + 1]; seg_l = l[s:ep + 1]
                    wo = np.append(wo, o[s]); wh = np.append(wh, np.nanmax(seg_h))
                    wl = np.append(wl, np.nanmin(seg_l)); wc = np.append(wc, c[ep]); partial = True
                if len(wc) >= 5:
                    try:
                        for k, val in _mp_indicators(wc, wh, wl, tag).items():
                            if val is not None and not (isinstance(val, float) and np.isnan(val)):
                                all_factors[k][row_idx] = val
                    except Exception:
                        pass
                    if len(wc) >= max(25, NM[1] + 2):
                        try:
                            smc = smc_factors(wo, wh, wl, wc, N=NM[0], M=NM[1])
                            for k, val in smc.items():
                                if not np.isnan(val):
                                    all_factors[f'{k}_{tag}'][row_idx] = val
                        except Exception:
                            pass
                if tag in tmp:                                    # turnover_multiperiod(td 轴, 完成柱)
                    rdates, rval, nbars = tmp[tag]
                    bc2 = int(np.searchsorted(rdates, d_str, 'right')) - 1
                    if bc2 >= nbars - 1:
                        rec = rval[bc2 - nbars + 1:bc2 + 1]
                        tsuf = '_W' if tag == 'W' else '_M'
                        all_factors[f'turnover_mean{tsuf}'][row_idx] = np.nanmean(rec)
                        all_factors[f'turnover_std{tsuf}'][row_idx] = np.nanstd(rec, ddof=1) if len(rec) >= 2 else np.nan
        done += 1
        if done % 100 == 0:
            print(f'    按股进度 {done}/{len(uniq)}')

    for k, arr in all_factors.items():
        df[k] = arr
    df.index = _orig_idx   # 恢复原 index(_derive_chunked concat+sort_index 还原全局序)
    shown = ', '.join(f'{k}({pd.Series(v).notna().mean()*100:.0f}%)' for k, v in sorted(all_factors.items())[:8])
    print(f'    完成 {done}/{len(uniq)} 股 | +{len(all_factors)}因子: {shown} ...')
    return df


# ====== J类: 月线10月均线趋势 (需 tushare 网络) ======

# ====== hfq 月线持久化(derive_monthly_trend 用, 免 fork 重复 pro_bar 撞限频) ======
_MONTHLY_CACHE = {}


def _monthly_bulk_read(codes):
    """从 stock_monthly_hfq 批量读 {code: (dates, close)}。表空→{}。"""
    if not codes:
        return {}
    try:
        import pymysql
        from utils.db_manager import ValuationDB
        cfg = ValuationDB.MYSQL_CONFIG
        conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                               password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
        try:
            ph = ','.join(['%s'] * len(codes))
            df = pd.read_sql(f"SELECT stock_code,trade_date,close FROM stock_monthly_hfq WHERE stock_code IN ({ph})",
                             conn, params=list(codes))
        finally:
            conn.close()
    except Exception:
        return {}
    if df is None or df.empty:
        return {}
    out = {}
    df, splits = _bulk_group_split(df)
    sc = df['stock_code'].to_numpy(); dates = df['trade_date'].to_numpy(); cl = df['close'].to_numpy(float)
    for idx in splits:
        if not len(idx):
            continue
        out[str(sc[idx[0]])] = (dates[idx], cl[idx])
    return out


def _monthly_save(code, dates, close):
    """一只股 hfq 月线落盘 stock_monthly_hfq(ON DUP KEY, 建表幂等)。"""
    if dates is None or len(dates) == 0:
        return
    try:
        import pymysql
        from utils.db_manager import ValuationDB
        cfg = ValuationDB.MYSQL_CONFIG
        conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                               password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
        n = len(dates)
        rows = [(code, str(dates[i]), _f(close[i])) for i in range(n)]
        try:
            with conn.cursor() as cur:
                cur.execute("""CREATE TABLE IF NOT EXISTS stock_monthly_hfq (
                    stock_code VARCHAR(16) NOT NULL, trade_date CHAR(8) NOT NULL, close DOUBLE,
                    PRIMARY KEY (stock_code, trade_date)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
                sql = "INSERT INTO stock_monthly_hfq (stock_code,trade_date,close) VALUES (%s,%s,%s) ON DUPLICATE KEY UPDATE close=VALUES(close)"
                B = 2000
                for i in range(0, n, B):
                    cur.executemany(sql, rows[i:i + B])
            conn.commit()
        finally:
            conn.close()
    except Exception:
        pass


def prefetch_monthly(codes, max_workers=12):
    """并发预取 hfq 月线入 _MONTHLY_CACHE + 落盘 stock_monthly_hfq。本地优先, 缺失才 tushare。"""
    from concurrent.futures import ThreadPoolExecutor
    import time
    uniq = [c for c in dict.fromkeys(str(c) for c in codes) if c not in _MONTHLY_CACHE]
    if not uniq:
        return
    local = _monthly_bulk_read(uniq)
    for c, r in local.items():
        _MONTHLY_CACHE[c] = r
    todo = [c for c in uniq if _MONTHLY_CACHE.get(c, (None, None))[0] is None]
    if not todo:
        print(f'    hfq月线: 全本地命中 {len(local)}, 零 tushare')
        return
    print(f'    hfq月线: 本地 {len(local)}, 待 tushare {len(todo)} (落盘)...')
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())

    def _work(c):
        for attempt in range(3):
            try:
                d = ts.pro_bar(ts_code=c, freq='M', adj='hfq')
                if d is None or len(d) == 0:
                    _MONTHLY_CACHE[c] = (None, None)
                    return False
                d = d.sort_values('trade_date').drop_duplicates('trade_date').reset_index(drop=True)
                res = (d['trade_date'].astype(str).to_numpy(), d['close'].astype(float).to_numpy())
                _MONTHLY_CACHE[c] = res
                _monthly_save(c, res[0], res[1])
                return True
            except Exception:
                time.sleep(1.0 * (attempt + 1))   # 限频退避
        _MONTHLY_CACHE[c] = (None, None)
        return False

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        list(ex.map(_work, todo))


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
    # 本地优先(免 fork 重复 pro_bar 撞限频): parent 已 prefetch_monthly 则 _MONTHLY_CACHE 热; 否则按块读本地+补缺
    if not any(c in _MONTHLY_CACHE and _MONTHLY_CACHE[c][0] is not None for c in unique_codes):
        prefetch_monthly(unique_codes)
    _mhit = sum(1 for c in unique_codes if c in _MONTHLY_CACHE and _MONTHLY_CACHE[c][0] is not None)
    print(f'    hfq 月线缓存: {_mhit}/{len(unique_codes)} 有数据')

    def get_monthly(code):
        if code not in _MONTHLY_CACHE:
            _MONTHLY_CACHE[code] = (None, None)
        return _MONTHLY_CACHE[code]

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

# 末尾闸门：剔除确认无信号的字段（IV=0 标记/常数/标识/定增参数/行业代码文本）
# placement 与 backtest 共用(run_derivation 调用); 拦住本脚本重新生成的无信号字段
DROP_FIELDS = {
    '市场指数_通过', '行业PE_通过', '个股PE_通过', 'DCF估值_通过',
    '修正PE估值_通过', '参数构造_通过', '蒙特卡洛_通过', '反向推算_通过',
    '子场景通过数', 'step1通过', 'step2通过', 'step3通过',
    '最新交易日', '行情_时间匹配', '财报年份', 'report_year',
    '市场_above_MA250',
    '溢价率', '溢价率下限', '锁定期', '融资金额', '无风险利率', 'Beta',
    '定增建议参与', '有效阈值数',
    'sw_l1_code', 'sw_l1_name', 'sw_l2_code', 'sw_l2_name', 'sw_l3_code', 'sw_l3_name',
    '营运资金变动_T', '营运资金变动_T1', '营运资金变动_T2', '营运资金变动_T3', '营运资金变动_T4',
    '营业利润率', '营收增长率', '行业PS', '净债务', '净资产负债表',
    '行业级别',
}


# ====== L类补: 行业指数技术面(idx_factor_pro PIT 切片, 87因子) ======

def derive_index_factor_features(df):
    """从 index_factor_pro PIT 切片(股票→行业指数, ≤报价日)产行业指数技术特征(87 因子)。
    data/refresh_industry_daily.ingest_idx_factor_pro 落表; 本函数 PIT 切片做回溯特征。"""
    print('\n  L类补: 行业指数技术面(idx_factor_pro PIT)...')
    if '股票代码' not in df.columns or '报价日' not in df.columns:
        return df
    s2i = _load_stock_to_idx()
    df['_行业idx'] = df['股票代码'].astype(str).map(s2i)
    unique_idx = [ic for ic in df['_行业idx'].dropna().unique() if ic]
    if not unique_idx:
        print('    无行业映射, 跳过'); return df
    import pymysql
    from utils.db_manager import ValuationDB
    conn = pymysql.connect(**ValuationDB.MYSQL_CONFIG)
    factor_cols = None; all_factors = {}; n_hit = 0
    for ic in unique_idx:
        try:
            idf = pd.read_sql("SELECT * FROM index_factor_pro WHERE index_code=%s ORDER BY trade_date", conn, params=(str(ic),))
        except Exception:
            continue
        if idf is None or len(idf) == 0:
            continue
        if factor_cols is None:
            factor_cols = [c for c in idf.columns if c not in ('index_code', 'trade_date')]
            for c in factor_cols:
                all_factors[f'行业idx_{c}'] = np.full(len(df), np.nan)
            print(f'    {len(factor_cols)} 因子, {len(unique_idx)} 行业指数')
        idf['td_int'] = pd.to_numeric(idf['trade_date'], errors='coerce').fillna(0).astype('int64')
        idf = idf.sort_values('td_int').reset_index(drop=True)
        td = idf['td_int'].to_numpy()
        mask = (df['_行业idx'] == ic).to_numpy()
        sub_dates = np.array([int(float(d)) for d in df.loc[mask, '报价日'].astype(str).to_numpy()])
        pos = np.searchsorted(td, sub_dates, 'right') - 1
        valid = pos >= 0
        if not valid.any():
            continue
        sub_indices = np.where(mask)[0][valid]
        for c in factor_cols:
            vals = idf[c].to_numpy()
            all_factors[f'行业idx_{c}'][sub_indices] = vals[pos[valid]]
        n_hit += len(sub_indices)
    conn.close()
    df = df.drop(columns=['_行业idx'])
    if factor_cols:
        for col, arr in all_factors.items():
            df[col] = arr
        print(f'    +{len(factor_cols)}特征(行业指数技术面, n_hit{n_hit})')
    return df


# ====== ProcessPool 分块并行: 3 个逐行派生(strategy/alpha_beta/monthly_trend)串行~14h → 并行~2h ======
# 因子逻辑零改动: <1500股 parent preload→fork 继承热缓存(只读); ≥1500股 build_backtest_panel 清父缓存,
# 各子进程 prefetch 自块(读本地DB)。env PARA_WORKERS 可覆盖。
_PARA_WORKERS = min(int(os.environ.get('PARA_WORKERS', '3')), max(1, (os.cpu_count() or 4) - 1))   # 8GB Mac+VSCode 控内存


def _derive_chunked(df, fn):
    """按行均分 _PARA_WORKERS 块 → fork Pool 各跑 fn(unchanged) → concat 回原序。
    fn 为模块级函数(fork 继承 + 按 ref pickle)。行太少(<200/块)则串行退回。"""
    import multiprocessing as mp
    n = min(_PARA_WORKERS, max(1, len(df) // 200))
    if n <= 1:
        return fn(df)
    chunks = [df.iloc[i::n].copy() for i in range(n)]   # 交错分块(跨股, 负载均衡)
    print(f'    [{fn.__name__}] ProcessPool×{n} 并行({len(df)}行, 每块~{len(chunks[0])})...')
    with mp.get_context('fork').Pool(n) as p:
        parts = p.map(fn, chunks)
    return pd.concat(parts).sort_index()


def run_derivation(df, skip_placement=False, run_stage2=True):
    """可重用派生核心: base df(含 股票代码/报价日 + 基列) → Stage1+Stage2 派生 → 清洗。

    **一套派生逻辑, placement 与 backtest/全A 共用**(后者 skip_placement=True,
    因 manifest 无定增原料)。所有 derive_* 防御式: 缺输入列自动跳过, 不崩。
    不落盘、不写 DB 快照(由调用方决定)——便于 backtest panel 构建器内联调用。
    """
    # ====== Stage 1: Parquet-only ======
    print('\n' + '='*60)
    print('Stage 1: Parquet-only 衍生特征')
    print('='*60)
    df = derive_fcf_growth_rates(df)
    df = derive_fcf_cross_metrics(df)
    df = derive_financial_score_deltas(df)
    df = derive_peer_valuation(df)   # 个股PE/PS/PB/市值 PIT + 同行截面均值 → 喂 _vs_*
    df = derive_valuation_relative(df)
    df = derive_market_stats_from_ohlcv(df)   # 行情统计基列(MA/波动率/年化收益..., 从OHLCV PIT) → 喂 derive_market_momentum
    df = derive_market_momentum(df)
    if not skip_placement:
        df = derive_placement_structure(df)

    # ====== Stage 2: DB/tushare-dependent (各派生独立 try, 单点失败不影响其余) ======
    if run_stage2:
        print('\n' + '='*60)
        print('Stage 2: DB-dependent 衍生特征')
        print('='*60)
        # parent preload 按 universe 大小自适应: 小(<1500股, ~0.4GB安全) → 预加载 fork 共享(免子进程冗余取);
        # 大(≥1500股) → 跳过(8GB Mac OOM), 各 fork 子进程按块读本地(stock_qfq_daily/stock_daily_basic)。
        try:
            _nu = df['股票代码'].astype(str).nunique()
            if _nu < 1500:
                _all_codes = df['股票代码'].astype(str).unique().tolist()
                print(f'  预热缓存(OHLCV+daily_basic+monthly)小universe({_nu}股<1500): fork 子进程共享, 免冗余取...')
                prefetch_ohlcv(_all_codes)
                prefetch_daily_basic(_all_codes)
                prefetch_monthly(_all_codes)
            else:
                print(f'  universe {_nu}股≥1500: 跳过 parent preload(8GB OOM), 各 fork 子进程按块读本地')
        except Exception as e:
            print(f'  ⚠️ 缓存预热失败(子进程将各自取): {e}')
        try:
            df = derive_industry_valuation_growth(df)
            df = derive_market_index_features(df)
        except Exception as e:
            print(f'  ⚠️ DB特征失败: {e}')
        try:
            df = derive_pb_vs_industry_pit(df)
        except Exception as e:
            print(f'  ⚠️ PB PIT 行业口径失败: {e}(保留原 peer 口径)')
        # ── 3 个逐行派生改 ProcessPool 分块并行(因子逻辑不动) ──
        try:
            df = _derive_chunked(df, derive_strategy_signals)
        except Exception as e:
            import traceback
            print(f'  ⚠️ 策略信号(H/I类)失败: {e}')
            traceback.print_exc()
        try:
            df = _derive_chunked(df, derive_alpha_beta_factors)
        except Exception as e:
            import traceback
            print(f'  ⚠️ 因子引擎(L类)失败: {e}')
            traceback.print_exc()
        try:
            df = _derive_chunked(df, derive_monthly_trend)
        except Exception as e:
            import traceback
            print(f'  ⚠️ 月线趋势(J类)失败: {e}')
            traceback.print_exc()
        try:
            df = derive_market_trend(df)
        except Exception as e:
            import traceback
            print(f'  ⚠️ 大盘趋势(K类)失败: {e}')
            traceback.print_exc()

        try:
            df = derive_index_factor_features(df)
        except Exception as e:
            print(f'  ⚠️ 行业指数技术面失败: {e}')

    # ====== 清理 ======
    df = df.replace([np.inf, -np.inf], np.nan)
    drop_present = [c for c in DROP_FIELDS if c in df.columns]
    if drop_present:
        df = df.drop(columns=drop_present)
        print(f'  末尾剔除无信号字段: {len(drop_present)} 个')
    return df


def main():
    parser = argparse.ArgumentParser(description='定增特征衍生 - 从基础特征派生高级特征')
    parser.add_argument('input_path', nargs='?',
                        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'features.parquet'),
                        help='输入特征文件 (parquet或CSV)')
    parser.add_argument('--output', default=None,
                        help='输出路径 (默认: 同目录下 features_derived.parquet)')
    parser.add_argument('--no-db', action='store_true', help='跳过需MySQL的特征(F+G类)')
    parser.add_argument('--skip-placement', action='store_true', help='跳过定增结构衍生(全A/回测 manifest 无定增原料时用)')
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

    # 派生(可重用核心 run_derivation; placement 与 backtest 共用一套派生逻辑)
    df = run_derivation(df, skip_placement=args.skip_placement, run_stage2=not args.no_db)

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
