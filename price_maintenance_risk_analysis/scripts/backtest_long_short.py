#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""全 A 多空组合回测 + IC/ICIR(量化核心验证)。

重心从 LOYO AUC 转到组合层面: 月度调仓全 A、7m 前瞻(=解禁期=模型标签口径)、
top/bottom 10% 等权多空 → IC/ICIR + L-S NAV(年化+maxDD) + 12 日历月分组(非重叠干净夏普)。

复用(不重写):
  - derive_features.derive_alpha_beta_factors  (价量/行业/beta, PIT ≤报价日 切片)
  - export_features.load_financial_ratios      (财务比率, PIT ann_date≤D)
  - predict_profitability.score_sc             (SC 打分, 模块级)
  - compute_labels.bench_return / add_months   (7m 前瞻收益)
  - fetch_universe                             (全 A + PIT 成员)

pilot-first: --sample 500 先验信号是否衰减, 过决策门再 --sample 0 全量。

用法:
  python scripts/backtest_long_short.py --horizon 7 --sample 500 --years 2010-2025
  python scripts/backtest_long_short.py --horizon 7 --sample 0    # 全 A
"""
import argparse
import os
import sys
import pickle
import calendar
from datetime import datetime

import numpy as np
import pandas as pd
from scipy.stats import spearmanr

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'ml_training'))
from predict_profitability import score_sc
from db_model_store import load_predict_bundle, get_model_meta
from report_horizon import latest_gray_sc
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from samples.fetch_universe import fetch_stock_basic, fetch_namechange, in_universe_at
from data_pipeline.compute_labels import add_months, bench_return, build_series, _nearest


# ─────────────── 月末交易日序列 ───────────────
def month_end_dates(year_start, year_end, tushare_calendar=None):
    """生成 yyyymmdd 月末交易日列表(用 tushare trade_cal, 无则取月末次日历)。"""
    if tushare_calendar is not None:
        cal = tushare_calendar
        out = []
        for y in range(year_start, year_end + 1):
            for m in range(1, 13):
                days = cal[(cal['year'] == y) & (cal['month'] == m) & (cal['is_open'] == 1)]
                if not days.empty:
                    out.append(days['date'].max())
        return sorted(out)
    # 退化: 月末日历日
    out = []
    for y in range(year_start, year_end + 1):
        for m in range(1, 13):
            out.append(f'{y}{m:02d}{calendar.monthrange(y, m)[1]:02d}')
    return out


# ─────────────── 特征组装(单截面 date) ───────────────
def compute_features(codes, date_yyyymmdd, model_feats):
    """为截面 date 上的 codes 算模型特征, 返回 DataFrame(index=codes, cols=model_feats)。
    复用 derive_alpha_beta_factors(价量/行业/beta) + load_financial_ratios(财务比率, PIT)。
    FCF/总分/北向 走独立 PIT 查询; SUE/PB_vs同行 暂 median 填(标注, 待精细化)。"""
    n = len(codes)
    rows = pd.DataFrame(index=codes)
    rows['股票代码'] = codes
    rows['报价日'] = date_yyyymmdd
    # 1) 价量/行业/beta(factor_engine, PIT ≤报价日)
    try:
        from derive_features import derive_alpha_beta_factors
        pv = derive_alpha_beta_factors(rows.copy())
        for c in pv.columns:
            if c in model_feats:
                rows[c] = pv[c].values
    except Exception as e:
        print(f'    derive_alpha_beta_factors 夻败: {e}')
    # 1b) 行业估值增长(行业PE/PB_{60,120,250}d增长) — derive_industry_valuation_growth
    try:
        from derive_features import derive_industry_valuation_growth
        iv = derive_industry_valuation_growth(rows.copy())
        for c in iv.columns:
            if c in model_feats:
                rows[c] = iv[c].values
    except Exception as e:
        print(f'    derive_industry_valuation_growth 失败: {e}')
    # 2) 财务比率(PIT ann_date≤date)
    try:
        from export_features import load_financial_ratios
        keys = [(c, date_yyyymmdd) for c in codes]
        fr = load_financial_ratios(keys)
        for c in fr.columns:
            if c in model_feats:
                rows[c] = fr[c].values
    except Exception as e:
        print(f'    load_financial_ratios 失败: {e}')
    # 3) 5 个特殊特征(独立 PIT loader: FCF/总分/nb_hold/PB_vs同行 PIT; SUE 桩)
    try:
        from export_features import load_specials
        sp = load_specials(codes, date_yyyymmdd)
        for c in sp.columns:
            if c in model_feats:
                rows[c] = sp[c].reindex(rows.index).values
    except Exception as e:
        print(f'    load_specials 失败: {e}')
    return rows.reindex(columns=model_feats)


# ─────────────── 前瞻收益(7m, 单截面) ───────────────
_CLOSE_CACHE = {}


def fwd_returns(codes, date_yyyymmdd, months=7):
    """每只 code 从 date 起 months 月的前瞻收益(%)。复用 compute_labels.bench_return。
    优先读 derive_features._OHLCV_CACHE(run_derivation 已全量预取 qfq), 免逐股二次 pro_bar。
    **持有期内退市/长期停牌**(找不到 D+months 价)→ 按退市前最后收盘计亏(PIT 诚实, 不 NaN 丢弃),
    否则生存偏差漏到收益层(差生被清场 → 收益虚高)。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    try:
        from derive_features import _OHLCV_CACHE as _OHC   # 同一 dict 对象(引用共享)
    except Exception:
        _OHC = {}
    out = {}
    for c in codes:
        if c not in _CLOSE_CACHE:
            cached = _OHC.get(c)
            if cached and cached[1] is not None:
                dates, ohlcv = cached
                _CLOSE_CACHE[c] = build_series(dict(zip(dates.tolist(), ohlcv['close'].tolist())))
            else:
                try:
                    df = ts.pro_bar(ts_code=c, adj='qfq', fields='trade_date,close')
                    _CLOSE_CACHE[c] = build_series(dict(zip(df['trade_date'], df['close']))) if df is not None else []
                except Exception:
                    _CLOSE_CACHE[c] = []
        out[c] = _bench_with_delist(_CLOSE_CACHE[c], date_yyyymmdd, months)
    return out   # {code: 收益%}


def _bench_with_delist(series, date_yyyymmdd, months):
    """bench_return 找不到 D+months 价(持有期退市/长期停牌)→ 按退市前最后收盘计亏。"""
    r = bench_return(series, date_yyyymmdd, months)
    if r is not None or not series:
        return r
    try:
        t0 = datetime.strptime(date_yyyymmdd, '%Y%m%d').toordinal()
    except ValueError:
        return None
    c0 = _nearest(series, t0)
    if not c0 or c0[1] == 0:
        return None
    last = series[-1]                       # 全局最后交易日(退市股=退市整理期最后价)
    if last[0] <= t0:                       # t0 之后无交易 → 仍 NaN
        return None
    return (last[1] / c0[1] - 1) * 100


# ─────────────── 单截面: 打分 + 排序 + 多空 + IC ───────────────
def eval_cross_section(codes, date, sc_bundle, model_feats, months=7, q=0.10):
    """返回 dict: ic, long_ret, short_ret, ls_ret, n。"""
    if len(codes) < 50:
        return None
    feat = compute_features(codes, date, model_feats)
    proba, _ = score_sc(sc_bundle, feat)
    fwd = fwd_returns(codes, date, months)
    s = pd.Series({c: fwd.get(c) for c in codes}).dropna()
    p = pd.Series(proba, index=feat.index).reindex(s.index)
    s = s.dropna(); p = p.dropna(); common = s.index.intersection(p.index)
    s, p = s.loc[common], p.loc[common]
    if len(s) < 50:
        return None
    ic = float(spearmanr(p, s).correlation)
    k = max(1, int(len(s) * q))
    order = p.sort_values()
    short_ret = s.loc[order.index[:k]].mean()
    long_ret = s.loc[order.index[-k:]].mean()
    return {'date': date, 'ic': ic, 'long': float(long_ret), 'short': float(short_ret),
            'ls': float(long_ret - short_ret), 'n': len(s)}


# ─────────────── 指标聚合 ───────────────
def nav_metrics(ls_series):
    """ls_series: 月度 L-S 收益序列(%)。返回 年化(CAGR) + maxDD(基于累计 NAV)。"""
    if len(ls_series) == 0:
        return {}
    r = np.array([x / 100.0 for x in ls_series], float)
    nav = np.cumprod(1 + r)
    cagr = nav[-1] ** (12 / len(r)) - 1 if nav[-1] > 0 else nav[-1] ** (12 / len(r)) - 1
    peak = np.maximum.accumulate(nav)
    dd = (nav - peak) / peak
    return {'cagr': float(cagr), 'maxdd': float(dd.min()), 'final_nav': float(nav[-1]),
            'mean_monthly': float(r.mean() * 100), 'n': len(r)}


def group_sharpes(records, months=7):
    """12 日历月分组(同月跨年, 非重叠) → 每组 ls 收益的 mean/std/年化夏普/年化(非重叠CAGR)。"""
    df = pd.DataFrame(records)
    if df.empty:
        return {}
    df['month'] = df['date'].astype(str).str[4:6].astype(int)
    out = {}
    for m, g in df.groupby('month'):
        ls = g['ls'].values / 100.0
        if len(ls) < 2:
            continue
        mu, sd = ls.mean(), ls.std(ddof=1)
        # 7m 收益 → 年化夏普: 每年 12/7 个非重叠 7m 期
        sharpe = (mu / sd) * np.sqrt(12 / months) if sd > 0 else 0
        # 7m 均值收益 → 年化(复利): (1+mu)^(12/months)-1。非重叠口径, 无重叠NAV虚高。
        ann_ret = (1 + mu) ** (12 / months) - 1
        out[m] = {'n': len(ls), 'mean_ls': float(mu * 100), 'std': float(sd * 100),
                  'sharpe': float(sharpe), 'ann_ret': float(ann_ret * 100)}
    return out


def group_leg_returns(records, months=7):
    """12 日历月分组 → 多头(top10%)和最差组(bottom10%)各自的 7m 收益/年化, 分开看(非重叠)。
    最差组按"做多这部分会怎样"算(正=赚, 负=亏), 不做空头。"""
    df = pd.DataFrame(records)
    if df.empty:
        return {}, 0.0, 0.0
    df['month'] = df['date'].astype(str).str[4:6].astype(int)
    out = {}
    for m, g in df.groupby('month'):
        lo = g['long'].values / 100.0      # top10% 多头 7m 收益
        sh = g['short'].values / 100.0     # bottom10% 最差组 7m 收益(做多口径)
        if len(lo) < 1:
            continue
        ann_lo = (1 + lo.mean()) ** (12 / months) - 1
        ann_sh = (1 + sh.mean()) ** (12 / months) - 1
        out[m] = {'n': len(lo),
                  'long_7m': float(lo.mean() * 100), 'long_ann': float(ann_lo * 100),
                  'short_7m': float(sh.mean() * 100), 'short_ann': float(ann_sh * 100)}
    avg_long = np.mean([v['long_ann'] for v in out.values()]) if out else 0
    avg_short = np.mean([v['short_ann'] for v in out.values()]) if out else 0
    return out, avg_long, avg_short


def group_by_year(records, months=7):
    """按 issue 年份分组 → 每年 多头/最差组/L-S 各自 7m 均值收益(做多口径, 看年度 regime)。
    年份分组天然非重叠(同年各月仓位收益取均值); 不做年化复利(单年 7m 收益直接报)。"""
    df = pd.DataFrame(records)
    if df.empty:
        return {}
    df['year'] = df['date'].astype(str).str[:4]
    out = {}
    for y, g in df.groupby('year'):
        lo = g['long'].values / 100.0
        sh = g['short'].values / 100.0
        ls = g['ls'].values / 100.0
        ic = g['ic'].values
        out[y] = {'n': len(lo),
                  'long_7m': float(lo.mean() * 100), 'short_7m': float(sh.mean() * 100),
                  'ls_7m': float(ls.mean() * 100),
                  'ic': float(ic.mean()) if len(ic) else 0.0}
    return out


# ─────────────── 主流程(读 panel, 不再运行时算特征) ───────────────
def run(model_ver, panel_path, horizon, q=0.10):
    """读 features_backtest.parquet → 逐截面打分 + 多空 + IC。
    特征/收益已由 build_backtest_panel 预计算入 panel, 此处只 score_sc + 排序。"""
    bundle = pickle.loads(load_predict_bundle(model_ver)['lr_bundle'])
    model_feats = bundle['features']
    panel = pd.read_parquet(panel_path)
    panel['报价日'] = panel['报价日'].astype(str)
    ret_col = f'return_{horizon}m'
    print(f'模型 {model_ver} | {len(model_feats)}特征 | {horizon}m | '
          f'panel {len(panel)}行 {panel["报价日"].nunique()}截面')

    records = []
    for d, g in panel.groupby('报价日'):
        if len(g) < 50:
            continue
        proba, _ = score_sc(bundle, g[model_feats].copy())
        s = g[ret_col]
        valid = s.notna()
        if valid.sum() < 50:
            continue
        p = pd.Series(proba, index=g.index)[valid]
        s = s[valid]
        ic = float(spearmanr(p, s).correlation)
        k = max(1, int(len(s) * q))
        order = p.sort_values()
        short = float(s.loc[order.index[:k]].mean())
        long = float(s.loc[order.index[-k:]].mean())
        records.append({'date': d, 'ic': ic, 'long': long, 'short': short,
                        'ls': long - short, 'n': int(len(s))})
        print(f"  {d}: n={len(s):4d} IC={ic:+.3f} L={long:+.2f}% "
              f"S={short:+.2f}% L-S={long-short:+.2f}%")
    if not records:
        print('❌ 无有效截面(n<50)'); return

    df = pd.DataFrame(records)
    ic = df['ic'].values
    icir = ic.mean() / ic.std() if ic.std() > 0 else 0
    nav = nav_metrics(df['ls'].values)
    grp = group_sharpes(records, horizon)
    grp_sharpe = np.mean([g['sharpe'] for g in grp.values()]) if grp else 0
    grp_pos = sum(1 for g in grp.values() if g['sharpe'] > 0)
    grp_ann = np.mean([g['ann_ret'] for g in grp.values()]) if grp else 0   # 非重叠CAGR

    print('\n' + '=' * 70)
    print(f'IC/ICIR: IC mean={ic.mean():+.4f} std={ic.std():.4f} → ICIR={icir:+.3f}')
    print(f'月度 L-S NAV(重叠, 偏乐观): CAGR={nav.get("cagr",0)*100:+.2f}% maxDD={nav.get("maxdd",0)*100:+.2f}%')
    print(f'12 日历月分组(非重叠, 诚实): 均值夏普={grp_sharpe:+.3f} | 均值年化={grp_ann:+.2f}% | {grp_pos}/12 组夏普为正')
    if grp:
        print('  每组 L-S(月: n / 7m均值% / 年化% / 夏普):')
        for m in sorted(grp):
            g = grp[m]
            print(f'    {m:2d}月: n={g["n"]} mean7m={g["mean_ls"]:+.2f}% 年化={g["ann_ret"]:+.2f}% 夏普={g["sharpe"]:+.2f}')
    # 多头/最差组分开(做多口径, 非重叠年化)
    legs, avg_long, avg_short = group_leg_returns(records, horizon)
    if legs:
        print(f'\n多头(top10%) vs 最差组(bottom10%, 做多口径) — 非重叠年化均值:')
        print(f'  投最好的 → 年化 {avg_long:+.2f}% | 投最差的 → 年化 {avg_short:+.2f}% | 差距 {avg_long-avg_short:+.2f}pp')
        print('  每组(月: 多头7m% / 多头年化% | 最差组7m% / 最差组年化%):')
        for m in sorted(legs):
            lg = legs[m]
            print(f'    {m:2d}月: 多 {lg["long_7m"]:+.2f}%/{lg["long_ann"]:+.2f}% | 差 {lg["short_7m"]:+.2f}%/{lg["short_ann"]:+.2f}%')
    # 按 issue 年份分组(看年度 regime: 哪年模型有效 top>>bottom, 哪年失效 bottom≈top)
    byr = group_by_year(records, horizon)
    if byr:
        print(f'\n按 issue 年份分组(7m 持有收益均值, 做多口径) — 看年度 regime:')
        print('  年份: n截面 | 多头7m% | 最差组7m% | L-S(差)% | IC')
        for y in sorted(byr):
            r = byr[y]
            print(f'  {y}: n={r["n"]:2d} | 多 {r["long_7m"]:+6.2f}% | 差 {r["short_7m"]:+6.2f}% | L-S {r["ls_7m"]:+6.2f}% | IC {r["ic"]:+.3f}')
    print('=' * 70)
    print('决策门: ICIR>0.3 且 非重叠年化明显为正 且 ≥9/12 组夏普为正 → 进全量')

    tag = os.path.basename(panel_path).replace('features_', '').replace('.parquet', '')
    out = os.path.join(PKG, 'ml_training', 'output', f'backtest_ls_{horizon}m_{tag}.csv')
    df.to_csv(out, index=False)
    print(f'写出: {out}')


def main():
    ap = argparse.ArgumentParser(description='全 A 多空组合回测 + IC/ICIR(读 panel)')
    ap.add_argument('--horizon', default='7', help='模型期限(月), 默认7(须与 panel 一致)')
    ap.add_argument('--panel', default=os.path.join(PKG, 'ml_training', 'data', 'features_backtest.parquet'),
                    help='回测特征 panel(build_backtest_panel.py 产出)')
    args = ap.parse_args()
    ver = latest_gray_sc(args.horizon)
    run(ver, args.panel, int(args.horizon))


if __name__ == '__main__':
    main()
