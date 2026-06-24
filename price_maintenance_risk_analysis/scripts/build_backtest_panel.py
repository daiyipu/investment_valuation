#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""建回测特征 panel: 样本清单 → 复用 compute_features + fwd_returns → features_backtest.parquet。

把 backtest_long_short 逐截面的"算特征+算7m前瞻收益"提前跑一遍落盘, 回测改读 panel
(废运行时 compute_features/feature_loaders)。脊柱 = backtest_samples.parquet(零 placement 依赖)。

复用(不重写特征逻辑):
  - backtest_long_short.compute_features  (derive_alpha_beta + derive_industry + ratios + specials)
  - backtest_long_short.fwd_returns        (bench_return 7m 前瞻)
  - feature_loaders.prefetch_fcf_scores    (FCF/总分 全历史预取)

用法:
  python build_backtest_panel.py                                       # 用 data/backtest_samples.parquet
  python build_backtest_panel.py --horizon 7                           # 指定模型期限(取特征清单)
"""
import argparse
import os
import sys
import pickle

import pandas as pd

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))           # price_maintenance_risk_analysis/
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'ml_training'))
sys.path.insert(0, os.path.join(PKG, 'ml_training', 'pipeline'))
sys.path.insert(0, os.path.join(PKG, 'scripts'))

from predict_profitability import score_sc  # noqa: E402 (确保 bundle 可加载)
from db_model_store import load_predict_bundle  # noqa: E402
from report_horizon import latest_gray_sc  # noqa: E402
from backtest_long_short import compute_features, fwd_returns  # noqa: E402
from feature_loaders import prefetch_fcf_scores, prefetch_sue_timelines  # noqa: E402

DATA_DIR = os.path.join(PKG, 'ml_training', 'data')
SAMPLES_PARQ = os.path.join(DATA_DIR, 'backtest_samples.parquet')
OUT_PARQ = os.path.join(DATA_DIR, 'features_backtest.parquet')


def main():
    ap = argparse.ArgumentParser(description='建回测特征 panel(复用 compute_features)')
    ap.add_argument('--horizon', default='7', help='模型期限(取特征清单), 默认7')
    ap.add_argument('--samples', default=SAMPLES_PARQ, help='样本清单 parquet')
    ap.add_argument('--out', default=OUT_PARQ, help='输出 panel parquet')
    args = ap.parse_args()

    samples = pd.read_parquet(args.samples)
    samples['报价日'] = samples['报价日'].astype(str)
    print(f'样本清单: {len(samples)} 条, {samples["报价日"].nunique()} 截面, {samples["股票代码"].nunique()} 唯一股')

    ver = latest_gray_sc(args.horizon)
    bundle = pickle.loads(load_predict_bundle(ver)['lr_bundle'])
    model_feats = bundle['features']
    print(f'模型 {ver} | {len(model_feats)} 特征 | horizon={args.horizon}m')

    # 预取全 universe FCF+总分+SUE 全历史(后续 compute_features→load_specials 内存 PIT 切片)
    uniq_codes = samples['股票代码'].astype(str).unique().tolist()
    prefetch_fcf_scores(uniq_codes)
    prefetch_sue_timelines(uniq_codes)

    # 逐截面: compute_features(15特征) + fwd_returns(7m) → 落盘
    frames = []
    dates = sorted(samples['报价日'].unique())
    for i, d in enumerate(dates):
        codes = samples.loc[samples['报价日'] == d, '股票代码'].astype(str).tolist()
        if len(codes) < 50:
            continue
        feat = compute_features(codes, d, model_feats)
        feat.insert(0, '股票代码', feat.index.astype(str))   # compute_features 仅返 model_feats 列(索引=codes), 重建股票代码列
        feat['报价日'] = d
        ret = fwd_returns(codes, d, months=int(args.horizon))
        feat[f'return_{args.horizon}m'] = feat['股票代码'].map(ret)
        frames.append(feat)
        if (i + 1) % 12 == 0:
            covered = sum(len(f) for f in frames)
            print(f'  截面 {i+1}/{len(dates)} ({d}): 累计 {covered} 行')
    if not frames:
        print('❌ 无有效截面(n<50)'); return
    panel = pd.concat(frames, ignore_index=True)
    panel.to_parquet(args.out, index=False)
    cov = panel[model_feats].notna().mean().mul(100).round(1)
    print(f'\n✅ panel: {len(panel)} 行 × {len(model_feats)} 特征 → {args.out}')
    print(f'   return_{args.horizon}m 覆盖: {panel[f"return_{args.horizon}m"].notna().mean()*100:.1f}%')
    print('   特征覆盖率(%):')
    for f, c in cov.items():
        print(f'     {f}: {c}')


if __name__ == '__main__':
    main()
