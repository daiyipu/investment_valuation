#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""特征选择标准流水线 —— 全项目唯一入口，不再各脚本各搞一套。

═══════════════════════════════════════════════════════════════════
固定五步 (阈值见下方常量, 已定标; 改阈值 = 只改这一处):
  1. IV top-N      候选池       单变量区分力(信息值)
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
  from feature_selection import select_features, prune_by_lgb_importance
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
from train_scorecard import calc_iv_all_features, remove_correlated, filter_by_vif

# ─────────────── 固定阈值(定标, 勿轻易改) ───────────────
N_IV = 40          # 步1: IV 候选池大小
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


def select_features(Xtr, ytr, Xte,
                    n_iv=N_IV, psi_max=PSI_MAX, corr_max=CORR_MAX,
                    vif_max=VIF_MAX, min_feat=MIN_FEAT, verbose=False):
    """标准五步之步 1-4: IV→PSI→去相关→VIF。

    返回 (kept_features, detail_df)。
    detail_df 列: feature, iv, psi, kept_psi, kept_corr, kept_vif, selected。
    """
    iv_df = (calc_iv_all_features(Xtr, ytr)
             .sort_values('iv', ascending=False)
             .head(n_iv).reset_index(drop=True))
    iv_map = dict(zip(iv_df['feature'], iv_df['iv']))

    detail = {}
    for _, r in iv_df.iterrows():
        f = r['feature']
        detail[f] = {'feature': f, 'iv': round(float(r['iv']), 3),
                     'psi': round(calc_psi(Xtr[f].values, Xte[f].values), 3)}

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
    n_psi = int(detail_df['kept_psi'].sum())
    n_corr = int(detail_df['kept_corr'].sum())
    n_vif = int(detail_df['kept_vif'].sum())
    return (f'IV top{n_iv} → PSI {n_psi} → 去相关 {n_corr} '
            f'→ VIF {n_vif} → LGBM剪 {len(kept_after_lgb)}')
