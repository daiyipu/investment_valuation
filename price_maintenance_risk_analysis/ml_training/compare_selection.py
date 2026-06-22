#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
特征选择方法对照：IV流程 vs Lasso vs 双向逐步回归

在同一批候选字段(IV>=iv_min，WOE 变换后)上，对比三种选字段方式：
  A. 当前评分卡流程: IV >= iv_min → 去相关(|r|>0.7) → VIF<5 → Top-N
  B. Lasso (L1 LR)：L1 正则自动把无用字段系数压到 0
  C. 双向逐步回归：向前加 AIC 最低的 + 向后剔能降 AIC 的

输出：各自选中的字段数 / 字段 / 5折CV AUC，便于判断哪种最精简且不损效果。
"""

import os, sys, warnings
import numpy as np
import pandas as pd
warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
sys.path.insert(0, os.path.join(SCRIPT_DIR, 'pipeline'))   # 管线模块已移入 pipeline/

from train_models import prepare_features_full
from train_scorecard import calc_iv_all_features, woe_transform
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import StratifiedKFold, cross_val_score
from sklearn.preprocessing import StandardScaler


def cv_auc(X, y, feats):
    """对选定 WOE 特征跑 5 折 LR CV AUC。"""
    Xs = X[feats].replace([np.inf, -np.inf], np.nan).fillna(0)
    scaler = StandardScaler()
    Xs2 = scaler.fit_transform(Xs.values)
    m = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    auc = cross_val_score(m, Xs2, y, cv=cv, scoring='roc_auc')
    return auc.mean(), auc.std()


def lasso_select(X_woe, y, features):
    """L1 LR，用 CV 选正则强度 C，返回系数非零字段。"""
    from sklearn.linear_model import LogisticRegressionCV
    Xs = X_woe[features].replace([np.inf, -np.inf], np.nan).fillna(0)
    scaler = StandardScaler()
    Xs2 = scaler.fit_transform(Xs.values)
    # C 越小正则越强(砍得越狠)；扫一组 C 用 CV(AUC) 选最优
    model = LogisticRegressionCV(
        Cs=np.logspace(-3, 1, 20), penalty='l1', solver='liblinear',
        cv=5, scoring='roc_auc', max_iter=2000, random_state=42,
    )
    model.fit(Xs2, y)
    coef = model.coef_[0]
    selected = [f for f, c in zip(features, coef) if abs(c) > 1e-6]
    print(f'  [Lasso] 最优 C={model.C_[0]:.4f}, 非零系数字段: {len(selected)}/{len(features)}')
    return selected


def stepwise_aic(X_woe, y, features):
    """双向逐步回归(基于 AIC)：向前加使 AIC 最低的，向后剔能使 AIC 更低的。"""
    import statsmodels.api as sm
    Xall = X_woe[features].replace([np.inf, -np.inf], np.nan).fillna(0)

    def fit_aic(feats):
        if not feats:
            return 1e9
        try:
            Xc = sm.add_constant(Xall[feats].values)
            m = sm.Logit(y.values, Xc).fit(disp=0, maxiter=200)
            return m.aic
        except Exception:
            return 1e9

    selected = []
    remaining = list(features)
    cur_aic = fit_aic(selected)
    improved = True
    while improved:
        improved = False
        # —— 向前：加入使 AIC 最低的字段 ——
        best_aic, best_f = cur_aic, None
        for f in remaining:
            a = fit_aic(selected + [f])
            if a < best_aic - 1e-6:
                best_aic, best_f = a, f
        if best_f:
            selected.append(best_f)
            remaining.remove(best_f)
            cur_aic = best_aic
            improved = True
        # —— 向后：剔除能使 AIC 更低的字段 ——
        for f in list(selected):
            a = fit_aic([s for s in selected if s != f])
            if a < cur_aic - 1e-6:
                selected.remove(f)
                cur_aic = a
                improved = True
        if len(selected) > 40:  # 安全上限
            break
    print(f'  [逐步回归] 收敛 AIC={cur_aic:.1f}, 选出 {len(selected)} 个字段')
    return selected


def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument('features_path', help='features_derived.parquet')
    p.add_argument('--threshold', type=float, default=-10)
    p.add_argument('--n', type=int, default=12, help='IV流程 Top-N')
    p.add_argument('--iv-min', type=float, default=0.05)
    args = p.parse_args()

    df = pd.read_parquet(args.features_path)
    X, y = prepare_features_full(df, args.threshold)
    if X is None:
        sys.exit(1)

    # 统一排除清单 (与 train_scorecard.py 共用 feature_exclusions)
    from feature_exclusions import get_excluded_columns
    excl = get_excluded_columns(X.columns)
    if excl:
        X = X.drop(columns=excl)
        print(f'(排除泄漏/artifact/业务字段: {excl})')
    print(f'样本 {len(y)}, 总特征 {X.shape[1]}, 盈利占比 {y.mean()*100:.1f}%\n')

    # 1. IV >= iv_min 候选池(三种方法的共同起点)
    iv_df = calc_iv_all_features(X, y)
    pool = iv_df[iv_df['iv'] >= args.iv_min]['feature'].tolist()
    print(f'IV>={args.iv_min} 候选池: {len(pool)} 个\n')

    # 2. WOE 变换候选池
    X_woe, _ = woe_transform(X, y, pool)
    woe_feats = [f for f in pool if f in X_woe.columns and X_woe[f].notna().any()]
    print(f'WOE 变换成功: {len(woe_feats)} 个\n')

    results = {}

    # A. 当前 IV 流程 Top-N (作为基准)
    print('=== A. 当前 IV 流程 (Top-%d) ===' % args.n)
    top_iv = iv_df[iv_df['feature'].isin(woe_feats)].nlargest(args.n, 'iv')['feature'].tolist()
    auc, std = cv_auc(X_woe, y, top_iv)
    results['IV Top-%d' % args.n] = (len(top_iv), top_iv, auc, std)
    print(f'  字段数={len(top_iv)}, CV AUC={auc:.3f}±{std:.3f}\n')

    # B. Lasso
    print('=== B. Lasso (L1 LR) ===')
    las = lasso_select(X_woe, y, woe_feats)
    if las:
        auc, std = cv_auc(X_woe, y, las)
        results['Lasso'] = (len(las), las, auc, std)
        print(f'  CV AUC={auc:.3f}±{std:.3f}\n')
    else:
        print('  ⚠ Lasso 全砍光，正则过强\n')

    # C. 双向逐步回归
    print('=== C. 双向逐步回归 (AIC) ===')
    sw = stepwise_aic(X_woe, y, woe_feats)
    if sw:
        auc, std = cv_auc(X_woe, y, sw)
        results['逐步回归'] = (len(sw), sw, auc, std)
        print(f'  CV AUC={auc:.3f}±{std:.3f}\n')

    # 汇总
    print('=' * 70)
    print(f'{"方法":<16} {"字段数":>6} {"CV AUC":>12}')
    print('-' * 40)
    for name, (nfeat, feats, auc, std) in results.items():
        print(f'{name:<16} {nfeat:>6} {auc:>8.3f}±{std:.3f}')
    print('=' * 70)

    # 字段明细
    for name, (nfeat, feats, auc, std) in results.items():
        print(f'\n【{name}】{nfeat}个, AUC={auc:.3f}:')
        print('  ' + ', '.join(feats))


if __name__ == '__main__':
    main()
