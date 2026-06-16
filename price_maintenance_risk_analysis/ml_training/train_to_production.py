#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""标准"从特征到生产"流水线(共识特征规则, 2026-06-17 固化)。

为什么: 标准五步选特征跨折不稳(6 折常只共有个别特征), 部署用某折的"一次抽取"
会和验证偏离。共识特征(跨折反复出现的)更稳、方差更小、可解释 → 作生产特征集。

流程(加完新特征后跑这一条即可):
  1. 各 LOYO 折 select_features → 特征跨折频次
  2. 共识 = 频次 ≥ --min-folds 的特征(跨期稳定的根本面; 太少则自动放宽)
  3. 锁定共识特征 LOYO(LGB/LR/SC) → 与部署对齐的真实性能(mean±std)
  4. 全量训最终模型(--model sc|lgb), 入库 set_current, 打印评分卡(SC)

用法:
  python train_to_production.py <features_derived.parquet> --horizon 7 --kind gray \
      [--min-folds 3] [--model sc] [--set-current]
"""
import argparse
import os
import sys
import pickle
import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_methods import make_features
from feature_selection import select_features
from train_horizon_models import GRAY_CFG, build_label, _prep, _train
from eval_loyo import loyo_fixed, fit_woe, apply_woe
from db_model_store import save_model_meta
from model_registry import register_version
from train_scorecard_model import print_scorecard, WOE_FILL, run as run_sc


def derive_consensus(df, horizon, kind, min_folds):
    """各 LOYO 折 select_features(真实 held-out 年判 PSI) → 特征频次 → 共识(频次≥min_folds)。"""
    df = df.dropna(subset=['报价日']).reset_index(drop=True)
    df['_y'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000).astype('Int64')
    years = sorted(int(y) for y in df['_y'].dropna().unique())
    from collections import Counter
    freq = Counter()
    ret = f'{horizon}个月涨跌幅'
    print(f'推导共识特征: {len(years)} 折各跑 select_features(真实 held-out 年判 PSI)...')
    for Y in years:
        dtr = df[df['_y'] != Y].drop(columns=['_y'])
        dte = df[df['_y'] == Y].drop(columns=['_y'])
        lbl, _ = build_label(dtr, horizon, kind)
        if kind == 'gray':
            build_label(dte, horizon, kind)
        Xtr_raw, ytr, _ = make_features(dtr, label_col=lbl, ret_col=ret)
        Xte_raw, yte, _ = make_features(dte, label_col=lbl, ret_col=ret)
        if Xtr_raw is None or Xte_raw is None:
            continue
        Xtr, med = _prep(Xtr_raw)
        Xte = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr.columns)
        kept, _ = select_features(Xtr, ytr, Xte)
        for f in kept:
            freq[f] += 1
        print(f'  折{Y} 选 {len(kept)}: {kept}')
    # 频次表
    print(f'\n特征跨折频次(共 {len(years)} 折):')
    for f, c in freq.most_common():
        print(f'  {f:<22} {c}/{len(years)} {"★" if c >= min_folds else ""}')
    consensus = [f for f, c in freq.most_common() if c >= min_folds]
    # 太少则放宽到 min_folds-1 (至少 4 个)
    mf = min_folds
    while len(consensus) < 4 and mf > 1:
        mf -= 1
        consensus = [f for f, c in freq.most_common() if c >= mf]
        print(f'  ⚠ 共识<4, 放宽 min_folds→{mf}')
    print(f'\n共识特征(min_folds={mf}, {len(consensus)}个): {consensus}')
    return consensus


def deploy_lgb(features_path, horizon, kind, consensus, split_year, set_current):
    """LGB 部署: 全量训 LGB+LR(共识特征), 入库。"""
    df = pd.read_parquet(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    lbl, gcfg = build_label(df, horizon, kind)
    Xall_raw, yall, _ = make_features(df, label_col=lbl, ret_col=f'{horizon}个月涨跌幅')
    Xall, med = _prep(Xall_raw)
    feats = [f for f in consensus if f in Xall.columns]
    gbm, lr, sc = _train(Xall[feats], yall)
    # LGB 训练集概率 10 分位边界(部署后 predict 映射固定档位)
    train_proba = gbm.predict_proba(Xall[feats])[:, 1]
    proba_deciles = np.quantile(train_proba, np.linspace(0.1, 0.9, 9)).tolist()
    ver = f'v_lgb_{pd.Timestamp.now().strftime("%Y%m%d_%H%M")}_{horizon}m_{kind}_{len(feats)}feat'
    save_model_meta({'version': ver, 'label_config': f'{horizon}m_{kind}_lgb_consensus',
                     'kind': kind, 'horizon': horizon, 'gray_cfg': gcfg, 'features': feats,
                     'n_features': len(feats), 'medians': {f: float(med[f]) for f in feats},
                     'lgb_model': gbm.booster_.model_to_string(),
                     'lr_bundle': pickle.dumps({'model': lr, 'scaler': sc, 'features': feats,
                                                 'proba_deciles': proba_deciles}),
                     'metrics': {}, 'note': f'LGB 共识特征 {feats}',
                     'dataset_version': 'derived_20260616_2334_f35ba6f3_7m'})
    register_version('full', ver, ver, metrics={}, n_features=len(feats), threshold=-10,
                     n_samples=len(yall), positive_rate=float(yall.mean()),
                     files=['(in DB)'], note=f'LGB共识{feats}', set_current=set_current,
                     label_config=f'{horizon}m_{kind}_lgb_consensus')
    print(f'\n✅ LGB 入库 {ver} | {"已设生产" if set_current else "未切生产"}')


def main():
    ap = argparse.ArgumentParser(description='共识特征→生产 标准流水线')
    ap.add_argument('features_path')
    ap.add_argument('--horizon', type=int, default=7)
    ap.add_argument('--kind', choices=['thr', 'gray'], default='gray')
    ap.add_argument('--min-folds', type=int, default=3, help='共识阈值: 跨折出现≥此数(默认3/6)')
    ap.add_argument('--model', choices=['sc', 'lgb'], default='sc')
    ap.add_argument('--set-current', action='store_true')
    args = ap.parse_args()

    df = pd.read_parquet(args.features_path)
    consensus = derive_consensus(df, args.horizon, args.kind, args.min_folds)
    # 锁定共识特征 LOYO(对齐验证)
    loyo_fixed(args.features_path, args.horizon, args.kind, ','.join(consensus))
    # 部署
    if args.model == 'sc':
        print(f'\n--- 训 SC(共识特征) 并部署 ---')
        run_sc(args.features_path, args.horizon, args.kind, 2024, args.set_current, features=consensus)
    else:
        deploy_lgb(args.features_path, args.horizon, args.kind, consensus, 2024, args.set_current)


if __name__ == '__main__':
    main()
