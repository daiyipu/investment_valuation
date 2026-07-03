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

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # validate/→ml_training/→PKG
for _p in (PKG, os.path.join(PKG,'ml_training'), os.path.join(PKG,'ml_training','pipeline'), os.path.join(PKG,'scripts')):
    if _p not in sys.path: sys.path.insert(0, _p)
from deploy.predict_profitability import score_sc
from deploy.db_model_store import load_predict_bundle, get_model_meta
from report.report_horizon import latest_gray_sc
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from samples.fetch_universe import fetch_stock_basic, fetch_namechange, in_universe_at
from data.labels import add_months, bench_return, build_series, _nearest


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
        from features.derive_features import derive_alpha_beta_factors
        pv = derive_alpha_beta_factors(rows.copy())
        for c in pv.columns:
            if c in model_feats:
                rows[c] = pv[c].values
    except Exception as e:
        print(f'    derive_alpha_beta_factors 夻败: {e}')
    # 1b) 行业估值增长(行业PE/PB_{60,120,250}d增长) — derive_industry_valuation_growth
    try:
        from features.derive_features import derive_industry_valuation_growth
        iv = derive_industry_valuation_growth(rows.copy())
        for c in iv.columns:
            if c in model_feats:
                rows[c] = iv[c].values
    except Exception as e:
        print(f'    derive_industry_valuation_growth 失败: {e}')
    # 2) 财务比率(PIT ann_date≤date)
    try:
        from features.export_features import load_financial_ratios
        keys = [(c, date_yyyymmdd) for c in codes]
        fr = load_financial_ratios(keys)
        for c in fr.columns:
            if c in model_feats:
                rows[c] = fr[c].values
    except Exception as e:
        print(f'    load_financial_ratios 失败: {e}')
    # 3) 5 个特殊特征(独立 PIT loader: FCF/总分/nb_hold/PB_vs同行 PIT; SUE 桩)
    try:
        from features.export_features import load_specials
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
        from features.derive_features import _OHLCV_CACHE as _OHC   # 同一 dict 对象(引用共享)
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
def nav_metrics(ls_series, dates=None):
    """ls_series: 月度 L-S 收益序列(%)。返回 年化(CAGR) + maxDD(基于累计 NAV)。
    🔧 重要修复：clipping极端收益率，避免NAV计算异常
    dates: 可选日期序列，用于调试"""
    if len(ls_series) == 0:
        return {}

    # 🔧 修复：输入数据已经是百分比，转换为小数
    r = np.array(ls_series, float) / 100.0

    # 🔧 关键修复：限制极端收益率，避免NAV计算异常
    # 原因：某些月份样本量小（如5个样本），top/bottom分组可能只有1-2个股票
    # 如果这些股票有极端表现（如单月+300%或-145%），会导致NAV计算异常
    r_clipped = np.clip(r, -0.95, 10.0)  # 限制单月最大损失95%，最大收益1000%

    # 检测哪些值被clipping了
    clipped_mask = (r != r_clipped)
    if clipped_mask.any() and dates is not None:
        print(f"  [回撤调试] {clipped_mask.sum()}个极端值被clipping:")
        for i in np.where(clipped_mask)[0]:
            print(f"    {dates[i]}: {ls_series[i]:.2f}% → {r_clipped[i]*100:.2f}%")

    nav = np.cumprod(1 + r_clipped)
    cagr = nav[-1] ** (12 / len(r_clipped)) - 1 if nav[-1] > 0 else nav[-1] ** (12 / len(r_clipped)) - 1
    peak = np.maximum.accumulate(nav)
    dd = (nav - peak) / peak

    # 找到最大回撤的位置
    max_dd_idx = dd.argmin()
    if dates is not None:
        print(f"  [回撤调试] 最大回撤发生在: {dates[max_dd_idx]}, 回撤值: {dd[max_dd_idx]*100:.2f}%")
        print(f"  [回撤调试] 该点NAV: {nav[max_dd_idx]:.4f}, 峰值NAV: {peak[max_dd_idx]:.4f}")

    return {'cagr': float(cagr), 'maxdd': float(dd.min()), 'final_nav': float(nav[-1]),
            'mean_monthly': float(r_clipped.mean() * 100), 'n': len(r_clipped)}


def group_sharpes(records, months=7):
    """12 日历月分组(同月跨年, 非重叠) → 每组 ls 收益的 mean/std/年化夏普/年化(非重叠CAGR)。"""
    df = pd.DataFrame(records)
    if df.empty:
        return {}
    df['month'] = df['date'].astype(str).str[5:7].astype(int)  # 🔧 修复: YYYY-MM-DD格式提取MM
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
def run(model_ver, horizon, q=0.10, sample_type='all', min_samples=5):
    """直接从ml_features_wide表读取 → 逐截面打分 + 多空 + IC。
    不再需要parquet文件，统一从数据库宽表读取。
    min_samples: 最小截面样本数（定增数据默认5，全A数据默认50）

    🔧 重要修复：验证集应该排除定增样本，使用纯全A市场数据
    - 训练集：sample_type='placement' (定增数据)
    - 验证集：sample_type='fake_quote' (全A市场数据，排除定增)
    """
    import pymysql
    from utils.db_manager import ValuationDB

    bundle = pickle.loads(load_predict_bundle(model_ver)['lr_bundle'])
    model_feats = bundle['features']
    conn = pymysql.connect(**ValuationDB.MYSQL_CONFIG)

    # 构建查询SQL
    label_col = f'{horizon}个月涨跌幅'
    base_cols = ['股票代码', '报价日', label_col]

    # 检查特征列
    cur = conn.cursor()
    cur.execute('SHOW COLUMNS FROM ml_features_wide')
    all_cols = [c[0] for c in cur.fetchall()]

    available_feats = [f for f in model_feats if f in all_cols]
    if len(available_feats) < 5:
        print(f'❌ 可用特征不足({len(available_feats)} < 5)，无法继续')
        conn.close()
        return

    if len(available_feats) < len(model_feats):
        print(f'⚠️ 缺少{len(model_feats) - len(available_feats)}个特征，使用{len(available_feats)}个可用特征')

    # 构建特征列查询
    feature_cols = available_feats
    all_query_cols = base_cols + feature_cols
    cols_quoted = ', '.join([f'`{c}`' if c in feature_cols or c == label_col else c for c in all_query_cols])

    # 构建WHERE条件
    where_clauses = [f'`{label_col}` IS NOT NULL']

    # 🔧 重要修复：默认使用全A市场数据验证，排除定增样本
    if sample_type == 'all':
        print('  [验证修复] 检测到sample_type="all"，自动切换为fake_quote以排除定增样本')
        sample_type = 'fake_quote'
        min_samples = 50  # 全A市场数据最小样本量设为50

    if sample_type == 'placement':
        where_clauses.append('sample_type="placement"')
        print('  [验证修复] 使用定增数据验证（⚠️ 可能重复训练集）')
    elif sample_type == 'fake_quote':
        where_clauses.append('sample_type="fake_quote"')
        print('  [验证修复] ✅ 使用全A市场数据验证（排除定增，真实泛化测试）')

    sql = f'SELECT {cols_quoted} FROM ml_features_wide WHERE {" AND ".join(where_clauses)}'

    print(f'查询数据...')
    panel = pd.read_sql(sql, conn)
    conn.close()

    # 重命名标签列为return格式
    ret_col = f'return_{horizon}m'
    panel = panel.rename(columns={label_col: ret_col})

    panel['报价日'] = panel['报价日'].astype(str)

    print(f'模型 {model_ver} | {len(available_feats)}特征 | {horizon}m | '
          f'panel {len(panel)}行 {panel["报价日"].nunique()}截面 | 样本类型:{sample_type}')

    records = []
    for d, g in panel.groupby('报价日'):
        if len(g) < min_samples:
            continue

        # 确保使用可用特征
        available_feats = [f for f in model_feats if f in g.columns]
        if len(available_feats) < 5:
            continue

        try:
            proba, _ = score_sc(bundle, g[available_feats].copy())
        except Exception as e:
            print(f"  {d}: score_sc失败 {e}, 跳过")
            continue

        s = g[ret_col]
        valid = s.notna()
        if valid.sum() < min_samples:
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

    # 🔧 只使用大样本（>=50）计算组别级最大回撤
    df_large = df[df['n'] >= 50].copy()
    print(f"  [组别回撤] 总月份数: {len(df)}, 大样本月份数(>=50): {len(df_large)}")

    if len(df_large) >= 12:  # 至少需要12个月
        nav = nav_metrics(df_large['ls'].values, df_large['date'].values)
        print(f"  [组别回撤] 使用{len(df_large)}个大样本月份计算组别级最大回撤")
    else:
        print(f"  [组别回撤] 大样本不足，使用全部数据")
        nav = nav_metrics(df['ls'].values, df['date'].values)
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

    # 写出结果
    out = os.path.join(PKG, 'ml_training', 'output', f'backtest_ls_{horizon}m_all_{model_ver[:10]}.csv')
    df.to_csv(out, index=False)
    print(f'写出: {out}')


def main():
    ap = argparse.ArgumentParser(description='全 A 多空组合回测 + IC/ICIR(直接读ml_features_wide表)')
    ap.add_argument('--horizon', default='7', help='模型期限(月), 默认7')
    ap.add_argument('--sample-type', default='all', choices=['all', 'placement', 'fake_quote'],
                    help='样本类型: all(全量), placement(定增), fake_quote(全市场)')
    ap.add_argument('--model', help='指定模型版本(不指定则用最新生产模型)')
    ap.add_argument('--min-samples', type=int, default=5, help='最小截面样本数（定增5，全A 50）')
    args = ap.parse_args()

    if args.model:
        ver = args.model
    else:
        ver = latest_gray_sc(args.horizon)
        print(f'使用最新{args.horizon}m灰度SC模型: {ver}')

    run(ver, int(args.horizon), sample_type=args.sample_type, min_samples=args.min_samples)


if __name__ == '__main__':
    main()
