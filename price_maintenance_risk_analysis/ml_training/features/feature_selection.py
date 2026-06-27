#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""特征选择标准流水线 —— 全项目唯一入口，不再各脚本各搞一套。

═══════════════════════════════════════════════════════════════════
固定五步 (阈值见下方常量, 已定标; 改阈值 = 只改这一处):
  1. IV>0.02且top50 候选池      单变量区分力(IV下限+top50 双条件, 2026-06-20 收紧)
  2. PSI ≤ ψ_max   跨期稳定     train/test 分布漂移; 防 regime 走捷径
  3. |r| ≤ ρ_max   去相关       冗余特征(如 ROE vs ROE摊薄), 保留 IV 高者
  4. VIF < ν_max   多重共线     逐步剔 VIF 最大者
  5. LGBM imp > 0  树实际使用   多变量交叉验证; 训练后用 prune_by_lgb_importance 剪
═══════════════════════════════════════════════════════════════════

设计原则:
  - 阈值定标后固定, 不再随脚本/随实验变 → 结果可比、可复现。
  - 复用 train_scorecard 既有 remove_correlated / filter_by_vif / calc_iv_all_features。
  - 各训练脚本(如 train_horizon_models)只调 select_features + prune_by_lgb_importance。

用法:
  from features.feature_selection import select_features, prune_by_lgb_importance
  kept, detail = select_features(Xtr, ytr, Xte)   # 步 1-4
  gbm = train_lgb(Xtr[kept], ytr)                 # 训练
  final = prune_by_lgb_importance(gbm, kept)      # 步 5: 剔树未用特征
"""

import contextlib
import io
import os
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from train.train_scorecard import calc_iv_all_features, remove_correlated, filter_by_vif

# ─────────────── 固定阈值(定标, 勿轻易改) ───────────────
IV_MIN = 0.01      # 步1: IV 下限(2026-06-21 由 0.02 放宽; 前门放宽, 让 PSI/corr/VIF/共识/LOO 后段流程自己筛合理指标, 缓解 IV 单变量+train锚定偏见)
N_IV = 100         # 步1: IV top-N 硬上限(由 50 放宽到 100; 前门放宽, 防过拟合改靠后段共识频次+Part A+LOO, 不靠前门硬截断)
PSI_MAX = 0.25     # 步2: PSI 上界(<0.1稳, 0.1~0.25微移, >0.25弃)
CORR_MAX = 0.7     # 步3: 相关系数上界(|r|>此值视为冗余)
VIF_MAX = 5.0      # 步4: VIF 上界(<5 严格, <10 宽松)
MIN_FEAT = 10      # 步1-4 后最少保留特征数(防过滤过狠)


def calc_psi(train, test, bins=10):
    """Population Stability Index: train→test 分布漂移度。
    <0.1 稳, 0.1~0.25 微移, >0.25 显著漂移。用 train 等频分箱套 test。"""
    tr = pd.Series(train).replace([np.inf, -np.inf], np.nan).dropna()
    te = pd.Series(test).replace([np.inf, -np.inf], np.nan).dropna()
    if len(tr) < bins or len(te) < bins:
        return 0.0
    edges = np.unique(np.quantile(tr, np.linspace(0, 1, bins + 1)))
    if len(edges) < 3:                       # 近常数, 无法有效分箱 → 视为稳定
        return 0.0
    edges = edges.astype(float)
    edges[0], edges[-1] = -np.inf, np.inf
    tp = np.histogram(tr, bins=edges)[0] / len(tr)
    ep = np.histogram(te, bins=edges)[0] / len(te)
    eps = 1e-4
    tp, ep = np.clip(tp, eps, None), np.clip(ep, eps, None)
    return float(np.sum((ep - tp) * np.log(ep / tp)))


def select_features(Xtr, ytr, Xte, Xtr_raw=None, Xte_raw=None,
                    iv_min=IV_MIN, psi_max=PSI_MAX, corr_max=CORR_MAX,
                    vif_max=VIF_MAX, min_feat=MIN_FEAT, n_iv=N_IV, verbose=False):
    """标准五步之步 1-4: IV→PSI→去相关→VIF。

    Xtr/Xte: 填充后的训练/测试集(去相关/VIF 用)。
    Xtr_raw/Xte_raw: 未填充原始集(可选; IV/PSI 用它算, 避免 median 填充污染 IV ——
      缺失样本填成 median 会堆一箱稀释区分力, 见 iv-selection-bias memory)。不传则用 Xtr(旧行为)。

    返回 (kept_features, detail_df)。detail_df 列: feature, iv, psi, kept_psi, kept_corr, kept_vif, selected。
    """
    Xiv = Xtr_raw.reindex(columns=Xtr.columns) if Xtr_raw is not None else Xtr
    Xiv_te = Xte_raw.reindex(columns=Xtr.columns) if Xte_raw is not None else Xte
    iv_all = (calc_iv_all_features(Xiv, ytr)
              .sort_values('iv', ascending=False).reset_index(drop=True))
    iv_df = iv_all[iv_all['iv'] > iv_min].reset_index(drop=True)   # IV 下限进漏斗
    if n_iv:                                                       # 硬上限(默认 N_IV=50; 与 IV_MIN 双条件)
        iv_df = iv_df.head(int(n_iv))
    iv_map = dict(zip(iv_df['feature'], iv_df['iv']))

    detail = {}
    for _, r in iv_df.iterrows():
        f = r['feature']
        detail[f] = {'feature': f, 'iv': round(float(r['iv']), 3),
                     'psi': round(calc_psi(Xiv[f].values, Xiv_te[f].values), 3)}

    # 步2 PSI
    kept = [f for f in detail if detail[f]['psi'] <= psi_max]
    if len(kept) < min_feat:                 # PSI 过严回退: 取 PSI 最低 min_feat 个
        kept = sorted(detail, key=lambda f: detail[f]['psi'])[:min_feat]
    for f in detail:
        detail[f]['kept_psi'] = f in kept

    # 步3 去相关 (静音 train_scorecard 的 print)
    with contextlib.redirect_stdout(io.StringIO() if not verbose else sys.stdout):
        decorr = remove_correlated(Xtr, kept, iv_df, threshold=corr_max)
    for f in detail:
        detail[f]['kept_corr'] = f in decorr
    kept = decorr

    # 步4 VIF
    with contextlib.redirect_stdout(io.StringIO() if not verbose else sys.stdout):
        try:
            vif_kept = filter_by_vif(Xtr, kept, max_vif=vif_max)
        except Exception as e:
            print(f'  VIF 跳过(import/statsmodels): {e}', file=sys.stderr)
            vif_kept = list(kept)
    for f in detail:
        detail[f]['kept_vif'] = f in vif_kept
    kept = list(vif_kept)

    for f in detail:
        detail[f]['selected'] = f in kept

    df = pd.DataFrame(detail.values()).sort_values(
        ['selected', 'iv'], ascending=[False, False]).reset_index(drop=True)
    return kept, df


def prune_by_lgb_importance(model, features, min_imp=1):
    """步5: 剔除 LGBM 重要性 < min_imp 的特征(树几乎没用到)。
    返回保留特征列表。默认 min_imp=1(剔除完全 0 次分裂的)。"""
    imp = pd.Series(getattr(model, 'feature_importances_', []),
                    index=features).fillna(0)
    return [f for f in features if imp.get(f, 0) >= min_imp]


def pipeline_summary(kept_after_vif, kept_after_lgb, detail_df):
    """人读摘要: 各步留下多少。"""
    n_iv = len(detail_df)
    iv_floor = round(float(detail_df['iv'].min()), 3) if n_iv else 0
    n_psi = int(detail_df['kept_psi'].sum())
    n_corr = int(detail_df['kept_corr'].sum())
    n_vif = int(detail_df['kept_vif'].sum())
    return (f'IV>{iv_floor}({n_iv}个) → PSI {n_psi} → 去相关 {n_corr} '
            f'→ VIF {n_vif} → LGBM剪 {len(kept_after_lgb)}')


# Lasso 备选分支(多变量全样本选择; 见 iv-selection-bias memory)
LASSO_MIN_FEAT = 8     # Lasso 选数下限(|coef| top 补足)
LASSO_MAX_FEAT = 30    # Lasso 选数上限(正则太弱选太多时截断 top)


def select_features_lasso(Xtr_raw, ytr, min_feat=LASSO_MIN_FEAT, max_feat=LASSO_MAX_FEAT):
    """Lasso(L1) 多变量特征选择 —— 全样本 log-loss 优化, 不靠单变量 IV。

    为何: IV 单变量 + 非灰 + train锚定 三重偏见(见 iv-selection-bias memory)。
    Lasso 在全样本上多变量优化、学交互 + 负权重, LOYO 实测 pooledIC 优于 IV 流程。
    流程: 缺失预筛(>50% 缺弃) → 标准化 → L1 LogisticRegressionCV(CV 选 C) → 非0系数 → [min,max] 截断。

    Xtr_raw: 未填充原始训练集(内部 fillna median 给 Lasso)。
    返回 (kept_features, detail_df); detail_df 列: feature, abs_coef。
    """
    from sklearn.linear_model import LogisticRegressionCV
    from sklearn.preprocessing import StandardScaler
    miss = Xtr_raw.isna().mean()
    cols = [c for c in Xtr_raw.columns if miss[c] < 0.5]
    X = Xtr_raw[cols].fillna(Xtr_raw[cols].median()).replace([np.inf, -np.inf], 0)
    sc = StandardScaler().fit(X)
    lr = LogisticRegressionCV(penalty='l1', solver='saga', Cs=np.logspace(-4, -1, 8),
                              cv=5, max_iter=4000, n_jobs=-1, scoring='roc_auc')
    lr.fit(sc.transform(X), ytr.values)
    coefs = lr.coef_[0]
    order = np.argsort(-np.abs(coefs))
    sel = [c for c, cf in zip(cols, coefs) if abs(cf) > 1e-6]
    if len(sel) < min_feat:                              # L1 过狠 → |coef| top 补足
        sel = [cols[i] for i in order[:min_feat]]
    elif len(sel) > max_feat:                            # 正则太弱选太多 → 截断 top
        sel = [cols[i] for i in order[:max_feat]]
    detail = (pd.DataFrame({'feature': cols, 'abs_coef': np.abs(coefs)})
              .sort_values('abs_coef', ascending=False).reset_index(drop=True))
    return sel, detail
