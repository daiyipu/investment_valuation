#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lasso 特征选择 vs 当前 IV 流程 —— LOYO 7m 对照实验。

背景(2026-06-20 诊断结论):
  IV top50 单变量截断 + 只在非灰train 上算, 三重病灶:
    ① 锚定 train 期 regime(不代表未来; 基本面全train IV 0.007→近期 0.065 差9倍);
    ② 只看极端赢/输(非灰), 不代表含灰度的全体生产排序(动量 IV 高但全样本IC负);
    ③ 对缺失/分箱敏感(盈利能力_delta_1y IV 预处理不同 0.0006 vs 0.143 差200倍)。
  → 系统性误杀组合强/近期有效的基本面(4feat 被 IV top51 踢), 误选全体反向的动量。

本实验对照两条特征选择流程:
  A) IV 流程: select_features(canonical IV>0.02+top50→PSI→corr→VIF) —— 当前生产
  B) Lasso 流程: 缺失预筛 → L1 LogisticRegressionCV 选非0系数(多变量, 全样本log-loss优化)
两条最终都用标准 LR(C=1) 在各自特征集上重训, 唯一变量 = 特征集(选择方法)。
LOYO 16 折, 报 AUC/KS/IC(非灰) + pooled IC(全样本拼接, 最稳) + 跨年AUC std(稳定性) + 特征集差异。

用法: python ml_training/diagnose_lasso_vs_iv.py
"""
import os, sys
import numpy as np
import pandas as pd
from sklearn.linear_model import LogisticRegression, LogisticRegressionCV
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import roc_auc_score
from scipy.stats import spearmanr

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from train.train_horizon_models import _prep, GRAY_CFG
from validate_methods import make_features, calc_ks
from features.feature_selection import select_features

PARQUET = os.path.join(HERE, 'data', 'features_derived.parquet')
H = 7
BASIC = ['个股PB', '盈利能力_delta_1y', '行业年化收益_20d', 'capex_intensity']


def lasso_select(Xtr, ytr, min_feat=8, max_feat=30):
    """缺失预筛 → L1 LogisticRegressionCV 选非0系数(多变量, 不靠单变量IV)。
    Cs 偏强正则(logspace -4~-1)促稀疏; 选数限制在 [min_feat, max_feat]。"""
    miss = Xtr.isna().mean()
    cols = [c for c in Xtr.columns if miss[c] < 0.5]
    X = Xtr[cols].fillna(Xtr[cols].median()).replace([np.inf, -np.inf], 0)
    sc = StandardScaler().fit(X)
    lr = LogisticRegressionCV(penalty='l1', solver='saga', Cs=np.logspace(-4, -1, 8),
                              cv=5, max_iter=4000, n_jobs=-1, scoring='roc_auc')
    lr.fit(sc.transform(X), ytr.values)
    coefs = lr.coef_[0]
    order = np.argsort(-np.abs(coefs))
    sel = [c for c, cf in zip(cols, coefs) if abs(cf) > 1e-6]
    if len(sel) < min_feat:                          # L1 过狠 → 取 |coef| top
        sel = [cols[i] for i in order[:min_feat]]
    elif len(sel) > max_feat:                        # 正则太弱选太多 → 截断 top
        sel = [cols[i] for i in order[:max_feat]]
    return sel


def lr_pred(feats, Xtr, ytr, Xte):
    Xtr_f = Xtr[feats].apply(pd.to_numeric, errors='coerce').replace([np.inf, -np.inf], np.nan)
    Xte_f = Xte[feats].apply(pd.to_numeric, errors='coerce').replace([np.inf, -np.inf], np.nan)
    med = Xtr_f.median()
    Xtr_f = Xtr_f.fillna(med).fillna(0).values
    Xte_f = Xte_f.fillna(med).fillna(0).values
    sc = StandardScaler().fit(Xtr_f)
    lr = LogisticRegression(C=1.0, max_iter=1000).fit(sc.transform(Xtr_f), ytr.values)
    return lr.predict_proba(sc.transform(Xte_f))[:, 1]


def main():
    df = pd.read_parquet(PARQUET).dropna(subset=['报价日']).reset_index(drop=True)
    df['_y'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000)
    h = H; lo, hi = GRAY_CFG[h]
    years = sorted(int(x) for x in df['_y'].dropna().unique())
    res = {k: {'auc': [], 'ks': [], 'ic': [], 'nf': [], 'p': [], 'basic': []}
           for k in ['iv', 'lasso']}
    ret_pool = []
    print(f'Lasso vs IV 特征选择 | LOYO 7m gray | {len(years)}折\n', flush=True)
    print(f'{"年":<6}{"IV流程":>28}{"Lasso流程":>28}', flush=True)
    for Y in years:
        dtr = df[df['_y'] != Y].drop(columns=['_y']).copy()
        dte = df[df['_y'] == Y].drop(columns=['_y']).copy()
        for d in (dtr, dte):
            rr = pd.to_numeric(d[f'{h}个月涨跌幅'], errors='coerce')
            d['标签_dbg'] = np.where(rr > hi, 1, np.where(rr < lo, 0, np.nan))
        Xtr_raw, ytr, _ = make_features(dtr, label_col='标签_dbg', ret_col=f'{h}个月涨跌幅')
        Xte_raw, yte, ret_te = make_features(dte, label_col='标签_dbg', ret_col=f'{h}个月涨跌幅')
        if Xtr_raw is None or len(ytr) < 40 or yte.nunique() < 2 or len(yte) < 20:
            print(f'{Y:<6}(非灰不足)', flush=True); continue
        Xtr, med = _prep(Xtr_raw)
        Xte = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr.columns)
        kept_iv, _ = select_features(Xtr, ytr, Xte)
        feats_iv = kept_iv if len(kept_iv) >= 5 else list(Xtr.columns[:10])
        feats_la = lasso_select(Xtr, ytr)
        row = f'{Y:<6}'
        for tag, feats in [('iv', feats_iv), ('lasso', feats_la)]:
            p = lr_pred(feats, Xtr, ytr, Xte)
            try:
                auc = roc_auc_score(yte, p)
            except Exception:
                auc = np.nan
            ks = calc_ks(yte.values, p)
            v = np.isfinite(ret_te.values) & np.isfinite(p)
            ic = float(spearmanr(p[v], ret_te.values[v]).correlation) if v.sum() > 20 else np.nan
            res[tag]['auc'].append(auc); res[tag]['ks'].append(ks); res[tag]['ic'].append(ic)
            res[tag]['nf'].append(len(feats)); res[tag]['p'].append(p)
            res[tag]['basic'].append(sum(f in feats for f in BASIC))
            row += f'  {len(feats)}feat(b{sum(f in feats for f in BASIC)}) A{auc:.2f}'
        ret_pool.append(ret_te.values)
        print(row, flush=True)

    print('\n=== LOYO 汇总 ===', flush=True)
    print(f'{"流程":<10}{"均AUC":>8}{"AUC_std":>9}{"均KS":>8}{"均IC":>8}{"pooledIC":>10}{"均feat":>8}{"均basic":>9}', flush=True)
    ret_all = np.concatenate(ret_pool)
    for tag, nm in [('iv', 'IV流程'), ('lasso', 'Lasso')]:
        a = np.array(res[tag]['auc'])
        p_all = np.concatenate(res[tag]['p'])
        v = np.isfinite(ret_all) & np.isfinite(p_all)
        pic = spearmanr(p_all[v], ret_all[v]).correlation
        print(f'{nm:<10}{np.nanmean(a):>8.3f}{np.nanstd(a):>9.3f}{np.nanmean(res[tag]["ks"]):>8.3f}'
              f'{np.nanmean(res[tag]["ic"]):>8.3f}{pic:>+10.3f}{np.nanmean(res[tag]["nf"]):>8.1f}'
              f'{np.nanmean(res[tag]["basic"]):>9.1f}', flush=True)


if __name__ == '__main__':
    main()
