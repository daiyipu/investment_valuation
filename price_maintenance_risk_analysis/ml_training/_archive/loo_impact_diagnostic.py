#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""LOO 净贡献诊断 —— 封装式特征评估(用户提的"测评性指标")。

在 LOYO 结构下,对候选池里每个特征做"留一删除":测删除后验证集 mean AUC/KS 与 std 的变化。
- Δmean_AUC < 0 (删除使 mean 下降) → 该特征有价值, 应留。
- Δmean_AUC > 0 (删除使 mean 上升) → 该特征有害, 应删。
- Δstd < 0 (删除使方差下降) → 该特征增不稳, 倾向删。
同时报 KS。目标: mean↑ + std↓ (score = mean − λ·std)。

候选池 = 新26共识 ∪ {Part A 删的2个} ∪ (旧34里有、新26没有的)。
直接定位 AUC 降因: 哪些特征帮/拖 LOYO。

用法: python loo_impact_diagnostic.py <features_derived.parquet> [--horizon 7] [--kind gray] [--lambda 0.5]
"""
import argparse, contextlib, io, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
from validate.eval_loyo import loyo_fixed

# 新26共识(Part A 去重后) + Part A 删的2个 + 旧34里新26没有的
NEW26 = ['pvt_slope_20', '行业PB_60d增长', '行业PE_120d增长', '行业PE_250d增长', '行业PB_120d增长', 'ROC_M_3',
         '行业PB_250d增长', 'MACD_HIST', '营收_3年斜率', 'chip_avg_cost_dev', 'MACD_W_HIST', '三浪_gain',
         'PB_vs_行业', '净资产增长', '总营收增长', 'turnover_mean_20', 'nb_hold_ratio', 'vwap_dist', 'ROA',
         'MACD_M_DEA', 'MACD_M_HIST', '速动比率', '流动资产周转率', 'ROE摊薄', '已获利息倍数', 'turnover_mean_250']
PART_A_DROPPED = ['beta_mkt_60', 'turnover_now']
OLD34 = ['流动比率', '现金利息负债比', '产权比率', '净资产增长', '总营收增长', '研发费用率', '已获利息倍数', 'ROA',
         '盈利能力_delta_2y', 'pvt_slope_20', '应收账款周转率', 'ROE摊薄', '行业PB_60d增长', '行业PE_250d增长',
         '行业PB_120d增长', 'MACD_HIST', '行业PE_120d增长', '行业PB_250d增长', '营收_3年斜率', 'chip_avg_cost_dev',
         '流动资产周转率', 'MACD_W_HIST', '扣非净利增长', '三浪_gain', 'turnover_now', 'MACD_M_DEA', 'MACD_M_HIST',
         'vwap_dist', 'ROC_M_3', '净利增长', 'turnover_mean_20', 'ROE', 'MACD_W_DIF', '总资产净利率']


def quiet_loyo(features_path, horizon, kind, feats):
    """跑 loyo_fixed 但静音, 只返回 sc 的统计。"""
    with contextlib.redirect_stdout(io.StringIO()):
        out = loyo_fixed(features_path, horizon, kind, ','.join(feats))
    return out['sc'] if out and out.get('sc') else None


def main():
    ap = argparse.ArgumentParser(description='LOO 净贡献诊断 (LOYO 封装式特征评估)')
    ap.add_argument('features_path')
    ap.add_argument('--horizon', type=int, default=7)
    ap.add_argument('--kind', choices=['thr', 'gray'], default='gray')
    ap.add_argument('--lambda', dest='lam', type=float, default=0.5, help='score = mean_AUC - λ·std_AUC')
    args = ap.parse_args()

    # 候选池: 去重 + 只保留数据里存在的列
    pool = []
    for f in NEW26 + PART_A_DROPPED + [x for x in OLD34 if x not in NEW26]:
        if f not in pool:
            pool.append(f)
    cols = set(pd.read_parquet(args.features_path, columns=None).columns)
    pool = [f for f in pool if f in cols]
    print(f'候选池 {len(pool)} 特征: {pool}\n')

    # 基线: 全池 LOYO
    base = quiet_loyo(args.features_path, args.horizon, args.kind, pool)
    base_score = base['auc_mean'] - args.lam * base['auc_std']
    print(f'全集({len(pool)}) SC LOYO: AUC {base["auc_mean"]:.4f}±{base["auc_std"]:.4f} | '
          f'KS {base["ks_mean"]:.4f}±{base["ks_std"]:.4f} | score(λ={args.lam})={base_score:.4f}\n')
    print(f'{"特征":<20} {"删后AUC":<14} {"Δmean":<8} {"Δstd":<8} {"ΔKSmean":<8} {"判定":<6}')
    print('-' * 75)

    rows = []
    for f in pool:
        sub = [x for x in pool if x != f]
        r = quiet_loyo(args.features_path, args.horizon, args.kind, sub)
        d_auc = r['auc_mean'] - base['auc_mean']
        d_std = r['auc_std'] - base['auc_std']
        d_ks = r['ks_mean'] - base['ks_mean']
        # 判定: 删除后 mean 下降(Δ<0)=有价值留; 上升(Δ>0)=有害删
        if d_auc > 0.002:
            verdict = '删'
        elif d_auc < -0.002:
            verdict = '留'
        else:
            verdict = '中性'
        rows.append((f, r['auc_mean'], d_auc, d_std, d_ks, verdict))

    # 按 Δmean_AUC 降序(最有害的=最该删的在前面)
    for f, auc, d_auc, d_std, d_ks, verdict in sorted(rows, key=lambda x: -x[2]):
        print(f'{f:<20} {auc:.4f}        {d_auc:+.4f}  {d_std:+.4f}  {d_ks:+.4f}  {verdict}')

    print('\n说明: Δmean>0 = 删它使 AUC 升(有害→删); Δmean<0 = 删它使 AUC 降(有价值→留); '
          f'score = mean_AUC − {args.lam}·std_AUC')


if __name__ == '__main__':
    main()
