#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""建回测特征 panel = **全特征仓**(可移植派生全集) + 7m 前瞻标签。

脊柱 = backtest_samples.parquet(零 placement 依赖: 股票代码 + 月末报价日)。
**复用 derive_features 全部 11 个可移植派生函数**(均防御式: 缺输入列自动跳过, 不崩),
跳过 derive_placement_structure(定增结构, 全A 无原料)。批量装可移植基列后喂进派生链。

一次到位(~150 可移植特征): 重新选特征/重训全A 时直接读此 panel, 不必重算。
backtest_long_short.run() 仍只取其中模型 15 列打分; 多余列它忽略。

基列(批量装, PIT):
  - FCF_T/T1-4 (营收/营业利润/净利润_fcf/NOPAT/折旧/资本支出/FCF, year≤pit_year) ← historical_fcf
  - 总分/盈利能力/成长能力 × T/T-1/T-2/T-4 (T=pit_year)                   ← company_annual_scores
  - 财务比率 27 (ann_date≤报价日)                                          ← financial_indicators
派生(复用 derive_features, 全防御):
  Stage1: fcf_growth_rates / fcf_cross_metrics / financial_score_deltas /
          valuation_relative / market_momentum (skip placement_structure)
  Stage2: industry_valuation_growth / market_index_features / pb_vs_industry_pit /
          strategy_signals / alpha_beta_factors / monthly_trend / market_trend
  其中 fcf_growth_rates 出 FCF_加速, score_deltas 出 总分_delta_2y, pb_vs_industry_pit 出
  PB_vs_同行中位 → 模型 15 特征里的 3 个 special 由派生路径直接覆盖。
补(派生未覆盖): nb_hold_ratio / sue_beat (export_features PIT loader, 缓存)。
标签: return_7m (bench_return, _CLOSE_CACHE)。

用法:
  python build_backtest_panel.py                              # 全量(5302×192月)
  python build_backtest_panel.py --limit 2                    # smoke: 仅前2截面
  python build_backtest_panel.py --horizon 7 --skip-label     # 跳过标签(快速验特征)
"""
import argparse
import os
import sys

import numpy as np
import pandas as pd
import pymysql

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))           # price_maintenance_risk_analysis/
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'ml_training'))
sys.path.insert(0, os.path.join(PKG, 'ml_training', 'pipeline'))
sys.path.insert(0, os.path.join(PKG, 'scripts'))

from utils.db_manager import ValuationDB  # noqa: E402
from export_features import load_financial_ratios, load_fcf_bulk, _pit_year  # noqa: E402  (基列装载单源)
from derive_features import run_derivation, prefetch_ohlcv  # noqa: E402  (一套派生核心, placement/全A 共用)
from export_features import prefetch_sue_timelines, load_specials  # noqa: E402  (5 PIT loader 已并入 export_features)
from backtest_long_short import fwd_returns  # noqa: E402  (7m 前瞻, _CLOSE_CACHE)

DATA_DIR = os.path.join(PKG, 'ml_training', 'data')
SAMPLES_PARQ = os.path.join(DATA_DIR, 'backtest_samples.parquet')
OUT_PARQ = os.path.join(DATA_DIR, 'features_backtest.parquet')

# 评分英→中 + 年份后缀(与 SCORE_METRICS + SCORE_DELTAS 一致)
# 评分装载暂留此处(company_annual_scores 不属 load_db_features 范畴; FCF 已单源到 export_features)
SCORE_COL_MAP = {'total_score': '总分', 'profitability': '盈利能力', 'growth': '成长能力'}
SCORE_SUFFIXES = ['_T', '_T-1', '_T-2', '_T-3', '_T-4']   # T=pit_year, T-k=pit_year-k(全; 含历史丢失的T-3)


def _bulk_score_base(samples):
    """批量装 总分/盈利能力/成长能力 × T/T-1/T-2/T-4。PIT: report_year≤pit_year。"""
    codes = [str(c) for c in samples['股票代码'].unique()]
    cfg = ValuationDB.MYSQL_CONFIG
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'], password=cfg['password'],
                           database=cfg['database'], charset=cfg['charset'])
    ph = ','.join(['%s'] * len(codes))
    eng_cols = list(SCORE_COL_MAP.keys())
    df = pd.read_sql(f"SELECT stock_code,report_year,{','.join(eng_cols)} FROM company_annual_scores "
                     f"WHERE stock_code IN ({ph})", conn, params=codes)
    conn.close()
    grouped = {str(c): {int(ry): row for ry, row in g.set_index('report_year').iterrows()}
               for c, g in df.groupby('stock_code')}
    out_cols = [f'{cn}{s}' for cn in SCORE_COL_MAP.values() for s in SCORE_SUFFIXES]
    out = {c: [] for c in out_cols}
    n_hit = 0
    for code, idate in zip(samples['股票代码'], samples['报价日']):
        pit = _pit_year(idate)
        sc = grouped.get(str(code), {})
        if pit in sc:
            n_hit += 1
        for s in SCORE_SUFFIXES:
            yr = pit if s == '_T' else pit - int(s.split('-')[-1])   # 解析后缀: T→pit, T-k→pit-k(免依赖列表连续性)
            row = sc.get(yr)
            for eng, cn in SCORE_COL_MAP.items():
                out[f'{cn}{s}'].append(row[eng] if row is not None else np.nan)
    print(f'  评分基列: {len(out_cols)} 列 (pit年命中 {n_hit}/{len(samples)})')
    return pd.DataFrame(out, index=samples.index)


def _per_section_extras(df, horizon, skip_label):
    """逐截面补 specials(nb_hold_ratio + SUE 全系列 8 变体; FCF_加速/总分_delta_2y/PB_vs_同行中位 已由 derive 产) + return_{h}m 标签。
    load_specials 复用(缓存); fwd_returns 复用(_CLOSE_CACHE)。"""
    _SUE_COLS = ['sue_beat', 'sue_zscore', 'sue_pos_streak', 'sue_recency_d',
                 'sue_yoy_acc', 'sue_yoy_mean3', 'sue_yoy', 'sue_up_trend']
    _NB_COLS = ['nb_hold_ratio', 'nb_hold_chg_20d', 'nb_hold_chg_60d']
    for col in _NB_COLS + _SUE_COLS:
        df[col] = np.nan
    if not skip_label:
        df[f'return_{horizon}m'] = np.nan
    dates = sorted(df['报价日'].astype(str).unique())
    for i, d in enumerate(dates):
        idx = df.index[df['报价日'].astype(str) == d]
        codes = df.loc[idx, '股票代码'].astype(str).tolist()
        try:
            sp = load_specials(codes, d)
            for col in _NB_COLS + _SUE_COLS:
                if col in sp.columns:
                    df.loc[idx, col] = sp.reindex(index=codes)[col].values
        except Exception as e:
            print(f'    load_specials @ {d} 失败: {e}')
        if not skip_label:
            try:
                ret = fwd_returns(codes, d, months=horizon)
                df.loc[idx, f'return_{horizon}m'] = [ret.get(c) for c in codes]
            except Exception as e:
                print(f'    fwd_returns @ {d} 失败: {e}')
        if (i + 1) % 12 == 0:
            print(f'    截面 {i+1}/{len(dates)} ({d})')
    return df


def _filter_tradable(samples):
    """剔除"报价日当月无成交"的停牌行(用热的 _OHLCV_CACHE, 须先 prefetch_ohlcv)。
    这些股在报价日附近长期停牌 → 无 D 价 → return_NaN 污染; 投资者当天根本买不到, 应剔。
    判据: 报价日 D 的最近一笔成交(≤D)须落在 D 所在自然月内, 否则视为当月停牌。"""
    from derive_features import _OHLCV_CACHE
    keep = np.ones(len(samples), dtype=bool)
    for code, g in samples.groupby('股票代码'):
        cached = _OHLCV_CACHE.get(str(code))
        if cached is None or cached[0] is None:
            keep[g.index.values] = False
            continue
        sd = cached[0]
        ds = g['报价日'].astype(str).values
        pos = np.searchsorted(sd, ds, 'right') - 1
        valid = pos >= 0
        last = np.where(valid, sd[np.clip(pos, 0, len(sd) - 1)], '')
        month_start = np.array([d[:6] + '01' for d in ds])
        keep[g.index.values] = valid & (last >= month_start)
    return samples[keep].reset_index(drop=True)


def _build_batch(batch_df, horizon, skip_label):
    """单批端到端建 panel: prefetch→filter→基列→run_derivation→extras→清洗。
    8GB Mac: 分批把主 df 降到 ~120k 行避免 swap 拖垮逐行 Stage2。每批前清模块缓存(bound 内存)。"""
    import gc as _gc
    from derive_features import _OHLCV_CACHE as _OCV, _DAILY_BASIC_CACHE as _DBC, _MONTHLY_CACHE as _MC
    _OCV.clear(); _DBC.clear(); _MC.clear(); _gc.collect()
    bdf = batch_df.copy().reset_index(drop=True)   # _filter_tradable 的 keep[index] 须 0..n-1 连续
    bdf['报价日'] = bdf['报价日'].astype(str)
    prefetch_ohlcv(bdf['股票代码'].astype(str).unique())
    n0 = len(bdf)
    bdf = _filter_tradable(bdf)
    df = bdf[['股票代码', '报价日']].copy()
    df['股票代码'] = df['股票代码'].astype(str)
    fcf = load_fcf_bulk(list(zip(df['股票代码'].astype(str), df['报价日'].astype(str)))); fcf.index = df.index
    score = _bulk_score_base(df)
    ratios = load_financial_ratios(list(zip(df['股票代码'].astype(str), df['报价日'].astype(str)))); ratios.index = df.index
    df = pd.concat([df, fcf, score, ratios], axis=1)
    df = run_derivation(df, skip_placement=True)
    df = _per_section_extras(df, horizon, skip_label)
    df = df.replace([np.inf, -np.inf], np.nan)
    return df


def main():
    ap = argparse.ArgumentParser(description='建回测特征 panel = 全特征仓 + 标签(复用 derive 全派生)')
    ap.add_argument('--horizon', default='7', help='标签期限(月), 默认7')
    ap.add_argument('--samples', default=SAMPLES_PARQ, help='样本清单 parquet')
    ap.add_argument('--out', default=OUT_PARQ, help='输出 panel parquet')
    ap.add_argument('--limit', type=int, default=0, help='仅前 N 截面(0=全部, smoke 用)')
    ap.add_argument('--skip-label', action='store_true', help='跳过 return 标签(快速验特征)')
    ap.add_argument('--batch-size', type=int, default=1200,
                    help='按股分批大小(8GB Mac 单遍上限~1200-1500股; 默认1200自动分批; 0=单遍)')
    args = ap.parse_args()
    horizon = int(args.horizon)

    samples = pd.read_parquet(args.samples)
    samples['报价日'] = samples['报价日'].astype(str)
    if args.limit:
        keep = sorted(samples['报价日'].unique())[:args.limit]
        samples = samples[samples['报价日'].isin(keep)].reset_index(drop=True)
    print(f'样本: {len(samples)} 行, {samples["报价日"].nunique()} 截面, '
          f'{samples["股票代码"].nunique()} 唯一股')

    # ── 0. SUE 时间线全量预取一次(跨批复用; 后续 load_sue_beat 内存读) ──
    print('\n预取 SUE 时间线(全量, 跨批复用)...')
    try:
        prefetch_sue_timelines(samples['股票代码'].astype(str).unique().tolist())
    except Exception as e:
        print(f'  ⚠️ SUE 预取失败(非致命): {e}')

    # ── 按股分批构建(8GB Mac: 主 df 563k→每批 ~120k 行, 避免 swap 拖垮逐行 Stage2) ──
    import time as _time
    if args.batch_size and args.batch_size > 0:
        stocks = samples['股票代码'].astype(str).drop_duplicates().tolist()
        B = args.batch_size
        nb = (len(stocks) + B - 1) // B
        panels = []
        print(f'\n按股分批: {len(stocks)}股 / {nb}批 (每批 ~{B}股)')
        for bi, i in enumerate(range(0, len(stocks), B), 1):
            bstk = stocks[i:i + B]
            bdf = samples[samples['股票代码'].astype(str).isin(bstk)]
            t0 = _time.time()
            print(f'\n{"=" * 60}\n[batch {bi}/{nb}] {len(bstk)}股 / {len(bdf)}行...')
            try:
                panel = _build_batch(bdf, horizon, args.skip_label)
            except Exception as e:
                import traceback; traceback.print_exc()
                print(f'  ⚠️ batch {bi} 失败(跳过): {e}'); continue
            print(f'  batch {bi} ✅ {len(panel)}行×{len(panel.columns)}列, 用时 {_time.time() - t0:.0f}s')
            panels.append(panel)
        if not panels:
            print('❌ 所有批次失败'); return
        df = pd.concat(panels, ignore_index=True)
    else:
        df = _build_batch(samples, horizon, args.skip_label)

    # ── 清洗 + 落盘 ──
    df = df.replace([np.inf, -np.inf], np.nan)
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    df.to_parquet(args.out, index=False)
    print('\n' + '=' * 60)
    print(f'✅ 全特征仓: {len(df)} 行 × {len(df.columns)} 列 → {args.out}')
    if not args.skip_label:
        print(f'   return_{horizon}m 覆盖: {df[f"return_{horizon}m"].notna().mean()*100:.1f}%')
    # 关键特征覆盖率抽检
    chk = ['PB_vs_同行中位', 'FCF_加速', '总分_delta_2y', 'nb_hold_ratio', 'sue_beat',
           '净资产增长', '应收账款周转率', 'MACD_W_HIST', '行业PE_250d增长']
    print('   关键特征覆盖率:')
    for c in chk:
        if c in df.columns:
            print(f'     {c}: {df[c].notna().mean()*100:.1f}%')


if __name__ == '__main__':
    main()
