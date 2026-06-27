#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""标签下行阈值扫描 —— 双样本空间(灰度实战影响)。可迁移任意期限。

方法论(2026-06-20 固化, 见 README「标签设计方法论」): 标签 lose 阈值不靠拍脑袋, 用本扫描定。
  - 固定 win 阈值(默认 +10, 生产口径), 扫 lose 阈值 [0,-5,-10,-15,-20]。
  - 每个阈值 LOYO, 同时报【两套样本空间】:
    1) 非灰度 AUC/KS: 清晰赢/输(标签非灰)的区分度 = 训练口径。
       ⚠ 极端 lose 阈值(如 -20)会因只留极值样本而虚高(海市蜃楼)。
    2) 全样本 rank IC: Spearman(预测概率, 实际收益) 跨【所有】样本(含灰度) = 实战排序能力;
       + 灰度% (覆盖; 实战筛不掉, 越高→有效样本越少)。
  - 选 lose 阈值按【全样本 IC】(实战), 不是非灰 AUC。
    教训(1m): (−20,10) 非灰AUC 0.61 但全样本 IC≈0(灰度 75% 拖累) → 假象;
              真实甜点 (−5,10)/(−10,10) 全样本 IC 最高(~0.039)。

IC 量纲(月级 rank corr): <0.02 噪声 / 0.02-0.05 弱但可用 / 0.05-0.10 中 / >0.10 强。

用法:
  python ml_training/sweep_label.py <features_derived.parquet> --horizon 3
  python ml_training/sweep_label.py <parquet> --horizon 1 --win 10 --loses 0,-5,-10,-15,-20
"""
import argparse
import os
import sys

import numpy as np
import pandas as pd
from scipy.stats import spearmanr
from sklearn.metrics import roc_auc_score

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # train/→ml_training/→PKG
for _p in (PKG, os.path.join(PKG,'ml_training'), os.path.join(PKG,'ml_training','pipeline'), os.path.join(PKG,'scripts')):
    if _p not in sys.path: sys.path.insert(0, _p)
from validate_methods import make_features, calc_ks, eval_metrics
from feature_selection import select_features
from train.train_horizon_models import _prep, _train, _ret_col, _tag, _parse_horizon


def _fold(dtr, dte, lo, hi, h):
    ret_col = _ret_col(h)
    lbl = f'_sw_{lo}'
    for d in (dtr, dte):
        rr = pd.to_numeric(d[ret_col], errors='coerce')
        d[lbl] = np.where(rr > hi, 1, np.where(rr < lo, 0, np.nan))
    Xtr_raw, ytr, _ = make_features(dtr, label_col=lbl, ret_col=ret_col)
    Xte_raw, yte, _ = make_features(dte, label_col=lbl, ret_col=ret_col)
    if Xtr_raw is None or Xte_raw is None or len(ytr) < 40 or yte.nunique() < 2:
        return None
    Xtr, med = _prep(Xtr_raw)
    Xte = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr.columns)
    kept, _ = select_features(Xtr, ytr, Xte, Xtr_raw=Xtr_raw, Xte_raw=Xte_raw)
    if len(kept) < 2:
        kept = list(Xtr.columns[:10])
    gbm, _, _ = _train(Xtr[kept], ytr)
    # 非灰度 AUC/KS(eval_metrics: n<20 或单类→nan, nanmean 自动丢退化折, 防 2010 n=2 污染)
    p_ng = gbm.predict_proba(Xte[kept])[:, 1]
    _em = eval_metrics(yte.values, p_ng)
    auc, ks = _em['auc'], _em['ks']
    # 全样本打分(含灰度) → IC + 灰度%
    Xall = dte.reindex(columns=Xtr.columns).apply(pd.to_numeric, errors='coerce')
    Xall_s = Xall[kept].fillna({f: med.get(f, 0) for f in kept}).replace([np.inf, -np.inf], 0)
    p_all = gbm.predict_proba(Xall_s)[:, 1]
    rr = pd.to_numeric(dte[ret_col], errors='coerce')
    gray_frac = float(np.where((rr <= hi) & (rr >= lo), 1, 0).mean())
    v = np.isfinite(rr.values) & np.isfinite(p_all)
    ic = float(spearmanr(p_all[v], rr.values[v]).correlation) if v.sum() > 20 else np.nan
    return auc, ks, gray_frac, ic


def main():
    ap = argparse.ArgumentParser(description='标签下行阈值扫描(双样本空间, 灰度实战)')
    ap.add_argument('features_path')
    ap.add_argument('--horizon', type=_parse_horizon, default=1)
    ap.add_argument('--win', type=float, default=10, help='win 阈值(默认+10, 生产口径)')
    ap.add_argument('--loses', default='0,-5,-10,-15,-20', help='lose 阈值列表(逗号分隔)')
    args = ap.parse_args()
    h, hi = args.horizon, args.win
    loses = [float(x) for x in args.loses.split(',')]
    df = pd.read_parquet(args.features_path).dropna(subset=['报价日']).reset_index(drop=True)
    df['_year'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000).astype('Int64')
    years = sorted(int(y) for y in df['_year'].dropna().unique())
    print(f'{_tag(h)} 标签扫描(win=+{hi:g}%, lose 横扫) | 双样本空间(非灰AUC/KS + 全样本IC) | {len(years)}折LOYO\n', flush=True)
    print(f'{"(lose,win)":<12}{"灰度%":>7}{"非灰AUC":>12}{"非灰KS":>8}{"全样本IC":>13}', flush=True)
    print('-' * 56, flush=True)
    for lo in loses:
        aucs, kss, gfs, ics = [], [], [], []
        for Y in years:
            dtr = df[df['_year'] != Y].drop(columns=['_year']).copy()
            dte = df[df['_year'] == Y].drop(columns=['_year']).copy()
            try:
                r = _fold(dtr, dte, lo, hi, h)
            except Exception:
                r = None
            if r is None:
                continue
            aucs.append(r[0]); kss.append(r[1]); gfs.append(r[2]); ics.append(r[3])
        if not aucs:
            print(f'({lo:>4.0f},{hi:>2.0f})   样本不足', flush=True); continue
        print(f'({lo:>4.0f},{hi:>2.0f})   {np.nanmean(gfs)*100:>6.1f}%'
              f'{np.nanmean(aucs):>7.3f}±{np.nanstd(aucs):.2f}{np.nanmean(kss):>8.3f}'
              f'{np.nanmean(ics):>9.3f}±{np.nanstd(ics):.2f}', flush=True)


if __name__ == '__main__':
    main()
