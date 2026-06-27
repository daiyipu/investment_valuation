#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""IV 代表性诊断 —— 为何基本面没进 GREEN + IV 随样本空间变化(不具代表性)。

固化 2026-06-20 诊断结论(配合 diagnose_lasso_vs_iv.py 的改进实验):

A. 4 基本面(个股PB/盈利能力_delta_1y/行业年化收益_20d/capex_intensity)过不了 canonical IV 门
   → 这就是 GREEN 全动量、没基本面的根因(IV 门偏见, 非基本面无用):
   - 个股PB IV0.0443 排第51(差1名出 top50); 其余 IV 低(0.0006-0.036)/排名78-320
   - 即便过 IV 门, PSI 全 0.47-0.68(>>0.25) 仍被 PSI 门卡

B. IV 样本空间(train 非灰极端)三重不代表性:
   ① 时间锚定 regime: 行业年化 IV 全train0.029→近期0.107(4×); capex 0.007→0.065(9×)
   ② 只看极端非灰: 动量 IV 高(0.18-0.25)但全样本 IC 负; IV 高 ≠ 全体排序有用
   ③ 对缺失/分箱敏感: 盈利能力_delta_1y IV 预处理不同 0.0006 vs 0.143

结论: IV 单变量 + 非灰 + train锚定 三重偏见 → 误杀组合强/近期有效的基本面、误选全体反向动量。
改进: 多变量全样本选择 Lasso/Elastic Net(见 diagnose_lasso_vs_iv.py)。

用法: python ml_training/diagnose_iv_representativeness.py
"""
import os, sys
import numpy as np
import pandas as pd
from scipy.stats import spearmanr

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from train.train_scorecard import calc_iv_all_features
from train.train_horizon_models import GRAY_CFG, _prep
from validate.validate_methods import make_features
from features.feature_selection import calc_psi, IV_MIN, PSI_MAX, N_IV
from deploy.db_model_store import get_model_meta
from deploy.model_registry import get_current

PARQUET = os.path.join(HERE, 'data', 'features_derived.parquet')
H = 7
BASIC = ['个股PB', '盈利能力_delta_1y', '行业年化收益_20d', 'capex_intensity']


def main():
    df = pd.read_parquet(PARQUET).dropna(subset=['报价日']).reset_index(drop=True)
    df['_y'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000)
    h = H; lo, hi = GRAY_CFG[h]; rc = f'{h}个月涨跌幅'
    df['ret'] = pd.to_numeric(df[rc], errors='coerce')
    df['y'] = np.where(df['ret'] > hi, 1, np.where(df['ret'] < lo, 0, np.nan))

    gm = get_model_meta(get_current('full')) or {}
    gfeats = gm.get('features', [])
    want = list(dict.fromkeys(BASIC + gfeats))
    tr = df[df['_y'] <= 2024].copy()

    # 全池 IV 排名(train ≤2024 非灰)
    dtr = tr.copy(); dtr['_lbl'] = dtr['y']
    Xtr_raw, ytr, _ = make_features(dtr, label_col='_lbl', ret_col=rc)
    Xtr, med = _prep(Xtr_raw)
    dte = df[df['_y'] >= 2025].copy(); dte['_lbl'] = dte['y']
    Xte_raw, _, _ = make_features(dte, label_col='_lbl', ret_col=rc)
    Xte = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr.columns)
    iv_all = calc_iv_all_features(Xtr, ytr).sort_values('iv', ascending=False).reset_index(drop=True)
    iv_all['rank'] = range(1, len(iv_all) + 1)
    rankN = float(iv_all.iloc[N_IV - 1]['iv']) if len(iv_all) >= N_IV else float('nan')

    print(f'═══ A. 为何没进 GREEN: IV/排名/PSI (门槛 IV>{IV_MIN} 且 排名≤{N_IV}[边界{rankN:.4f}], PSI≤{PSI_MAX}) ═══\n')
    print(f'{"特征":<16}{"IV":>8}{"排名":>6}{"PSI":>7}{"IV门":>6}{"PSI门":>7}  组')
    for f in want:
        s = iv_all[iv_all['feature'] == f]
        iv = float(s['iv'].iloc[0]) if len(s) else np.nan
        rk = int(s['rank'].iloc[0]) if len(s) else -1
        psi = calc_psi(Xtr[f].values, Xte[f].values) if f in Xte.columns else np.nan
        piv = (iv > IV_MIN and rk <= N_IV)
        ppsi = psi <= PSI_MAX
        print(f'{f:<16}{iv:>8.4f}{rk:>6}{psi:>7.3f}{"✓" if piv else "✗":>6}{"✓" if ppsi else "✗":>7}  {"★BASIC" if f in BASIC else "GREEN"}')

    print(f'\n═══ B. IV 随样本空间变化(看代表性) ═══')
    print(f'  列: 非灰IV(全train≤24)=选择口径 | 全样本IC(≤24含灰) | 非灰IV(仅近期20-24)\n')

    def iv_of(sub):
        X = sub[want].apply(pd.to_numeric, errors='coerce')
        t = calc_iv_all_features(X, sub['y'])
        return dict(zip(t['feature'], t['iv']))
    ivA = iv_of(tr[tr['y'].notna()])
    rc_df = df[(df['_y'] >= 2020) & (df['_y'] <= 2024)].copy()
    ivC = iv_of(rc_df[rc_df['y'].notna()])
    tv = tr[tr['ret'].notna()]; icB = {}
    for f in want:
        a = pd.to_numeric(tv[f], errors='coerce'); v = np.isfinite(a) & np.isfinite(tv['ret'])
        icB[f] = spearmanr(a[v], tv['ret'][v]).correlation if v.sum() > 20 else np.nan
    print(f'{"特征":<16}{"IV全train":>10}{"全样本IC":>10}{"IV近期":>9}')
    for f in want:
        print(f'{f:<16}{ivA.get(f, np.nan):>10.4f}{icB.get(f, np.nan):>+10.3f}{ivC.get(f, np.nan):>9.4f}  {"★" if f in BASIC else ""}')


if __name__ == '__main__':
    main()
