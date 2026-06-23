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
from data_pipeline.fetch_universe import fetch_stock_basic, fetch_namechange, in_universe_at
from data_pipeline.compute_labels import add_months, bench_return, build_series


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
        print(f'    derive_alpha_beta_factors 失败: {e}')
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
    # 3) FCF_加速 / 总分_delta_2y / nb_hold_ratio 独立 PIT loader(略, 待接入)
    #    TODO: FCF(historical_fcf PIT) / 总分(company_annual_scores PIT) / nb_hold(hk_hold≤date)
    #    暂缺 → score_sc 会用训练 median 填(标注: 该特征在回测中暂无信号贡献)
    return rows[model_feats] if all(f in rows.columns for f in model_feats) else \
        rows.reindex(columns=model_feats)


# ─────────────── 前瞻收益(7m, 单截面) ───────────────
_CLOSE_CACHE = {}


def fwd_returns(codes, date_yyyymmdd, months=7):
    """每只 code 从 date 起 months 月的前瞻收益(%)。复用 compute_labels.bench_return。
    全量预取 close_map 缓存, 避免逐股 API。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    out = {}
    for c in codes:
        if c not in _CLOSE_CACHE:
            try:
                df = ts.pro_bar(ts_code=c, adj='qfq', fields='trade_date,close')
                _CLOSE_CACHE[c] = build_series(dict(zip(df['trade_date'], df['close']))) if df is not None else []
            except Exception:
                _CLOSE_CACHE[c] = []
        out[c] = bench_return(_CLOSE_CACHE[c], date_yyyymmdd, months)
    return out   # {code: 收益%}


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
    """12 日历月分组(同月跨年, 非重叠) → 每组 ls 收益的 mean/std/年化夏普。"""
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
        out[m] = {'n': len(ls), 'mean_ls': float(mu * 100), 'std': float(sd * 100), 'sharpe': float(sharpe)}
    return out


# ─────────────── 主流程 ───────────────
def run(model_ver, sample_n, year_start, year_end, months=7):
    bundle = pickle.loads(load_predict_bundle(model_ver)['lr_bundle'])
    model_feats = bundle['features']
    print(f'模型 {model_ver} | {len(model_feats)}特征 | {months}m前瞻 | sample={sample_n or "全A"}')

    stocks = fetch_stock_basic()
    nc = fetch_namechange()
    if sample_n:
        from data_pipeline.fetch_universe import sample_stratified
        picks = sample_stratified(stocks, sample_n)
        stocks = stocks[stocks['ts_code'].isin(picks)].reset_index(drop=True)
        print(f'pilot 分层抽样: {len(stocks)} 只')

    dates = month_end_dates(year_start, year_end)
    print(f'调仓: {len(dates)} 个月末截面 ({dates[0]}~{dates[-1]})\n')

    records = []
    for i, d in enumerate(dates):
        codes = [r['ts_code'] for _, r in stocks.iterrows()
                 if in_universe_at(r, d, nc, min_list_years=1)]
        res = eval_cross_section(codes, d, bundle, model_feats, months)
        if res:
            records.append(res)
            print(f"  {d}: n={res['n']:4d} IC={res['ic']:+.3f} L={res['long']:+.2f}% "
                  f"S={res['short']:+.2f}% L-S={res['ls']:+.2f}%")
    if not records:
        print('❌ 无有效截面'); return

    df = pd.DataFrame(records)
    ic = df['ic'].values
    icir = ic.mean() / ic.std() if ic.std() > 0 else 0
    nav = nav_metrics(df['ls'].values)
    grp = group_sharpes(records, months)
    grp_sharpe = np.mean([g['sharpe'] for g in grp.values()]) if grp else 0
    grp_pos = sum(1 for g in grp.values() if g['sharpe'] > 0)

    print('\n' + '=' * 70)
    print(f'IC/ICIR: IC mean={ic.mean():+.4f} std={ic.std():.4f} → ICIR={icir:+.3f}')
    print(f'月度 L-S NAV(重叠): CAGR={nav.get("cagr",0)*100:+.2f}% maxDD={nav.get("maxdd",0)*100:+.2f}%')
    print(f'12 日历月分组(非重叠): 均值夏普={grp_sharpe:+.3f} | {grp_pos}/12 组夏普为正')
    print('=' * 70)
    print('决策门: ICIR>0.3 且 L-S CAGR 明显为正 且 ≥9/12 组夏普为正 → 进全量')

    out = os.path.join(PKG, 'price_maintenance_risk_analysis', 'ml_training', 'output',
                       f'backtest_ls_{months}m_{"sample"+str(sample_n) if sample_n else "allA"}.csv')
    df.to_csv(out, index=False)
    print(f'写出: {out}')


def main():
    ap = argparse.ArgumentParser(description='全 A 多空组合回测 + IC/ICIR')
    ap.add_argument('--horizon', default=7, help='模型期限(月), 默认7')
    ap.add_argument('--sample', type=int, default=500, help='分层抽样数(0=全A)')
    ap.add_argument('--years', default='2010-2025')
    ap.add_argument('--months', type=int, default=7, help='前瞻/持仓月数(=解禁期)')
    args = ap.parse_args()
    ys, ye = map(int, args.years.split('-'))
    ver = latest_gray_sc(args.horizon)
    run(ver, args.sample, ys, ye, args.months)


if __name__ == '__main__':
    main()
