#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""训练 WOE 评分卡(SC)模型并入库; 可设为生产(current.full)。

标准五步选特征(用 split 判 PSI) → fit_woe → LR on WOE → 全量重训 → 入库。
bundle 存 ml_model_meta.lr_bundle({kind:'scorecard', woe_bins, lr_model, features, medians}),
lgb_model=NULL。predict 检测 kind=='scorecard' 走 WOE 打分路径(替代 LGB)。

用法: python train_scorecard_model.py <features_derived.parquet> --horizon 7 --kind gray [--set-current]
"""
import argparse
import os
import sys
import pickle
import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_methods import make_features, eval_metrics
from feature_selection import select_features, pipeline_summary, IV_MIN, PSI_MAX, CORR_MAX, VIF_MAX
from train_horizon_models import GRAY_CFG, build_label, _prep, _ret_col, _tag, _parse_horizon
from eval_loyo import fit_woe, apply_woe
from db_model_store import save_model_meta
from model_registry import register_version
from sklearn.linear_model import LogisticRegression

WOE_FILL = lambda X: X.replace([np.inf, -np.inf], np.nan).fillna(0)


def print_scorecard(features, woe_bins, lr):
    """打印评分卡: 每特征的 WOE 分箱(边界+woe) + LR 系数, 按|系数|降序。"""
    coefs = dict(zip(features, lr.coef_[0]))
    order = sorted(features, key=lambda f: abs(coefs[f]), reverse=True)
    print('\n' + '=' * 90)
    print(f'评分卡: {len(features)} 特征 | LR intercept={lr.intercept_[0]:.4f}')
    print('=' * 90)
    for f in order:
        c = coefs[f]
        print(f'\n■ {f}  (LR系数={c:+.4f})')
        if f in woe_bins:
            rights = woe_bins[f]['rights']; woes = woe_bins[f]['woes']
            for i, (r, w) in enumerate(zip(rights, woes)):
                edge = '−∞' if i == 0 else f'{rights[i-1]:.4g}'
                print(f'    分箱{i+1}: ({edge}, {r:.4g}]  → WOE={w:+.4f}')
        else:
            print('    (无分箱/常数)')


def run(features_path, horizon, kind, split_year, set_current, features=None, loyo_stats=None):
    df = pd.read_parquet(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    df['_y'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000).astype('Int64')
    dtr_s = df[df['_y'] <= split_year].drop(columns=['_y'])      # 选特征用(算 PSI)
    dte_s = df[df['_y'] >= split_year + 1].drop(columns=['_y'])
    lbl, gcfg = build_label(dtr_s, horizon, kind)
    if kind == 'gray':
        build_label(dte_s, horizon, kind)
    ret = _ret_col(horizon)

    # 特征: --features 锁定(共识特征, 跳过选择) 或 标准五步选
    Xtr_raw, ytr_s, _ = make_features(dtr_s, label_col=lbl, ret_col=ret)
    Xte_raw, yte_s, _ = make_features(dte_s, label_col=lbl, ret_col=ret)
    Xtr_s, med = _prep(Xtr_raw)
    Xte_s = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr_s.columns)
    if features:
        kept = [f for f in features if f in Xtr_s.columns]
        detail = None
        sel = f'locked({len(kept)}共识特征, 名单见 features 列)'
        print(f'⚠ 锁定特征(跳过select_features): {kept}')
    else:
        kept, detail = select_features(Xtr_s, ytr_s, Xte_s, Xtr_raw=Xtr_raw, Xte_raw=Xte_raw)
        sel = pipeline_summary(kept, kept, detail)
        print(f'选择漏斗: {sel} → 入模 {len(kept)} 特征: {kept}')

    # 全量重训(部署用全部数据): build_label on 全量 df
    lbl2, _ = build_label(df, horizon, kind)
    Xall_raw, yall, _ = make_features(df.drop(columns=['_y']), label_col=lbl2, ret_col=ret)
    Xall, _ = _prep(Xall_raw)
    Xall_w, wbins = fit_woe(Xall[kept], yall, kept)
    lr = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)
    lr.fit(WOE_FILL(Xall_w), yall)

    # 全量训练后自评: test 年已入训练 → 含泄漏偏高, 非真实泛化(仅拟合参考)。
    # 真泛化能力看 loyo_stats(由 train_to_production 跑 loyo_fixed 传入)。
    Xte_s_w = apply_woe(Xte_s[kept], kept, wbins)
    ate = eval_metrics(yte_s.values, lr.predict_proba(WOE_FILL(Xte_s_w))[:, 1])
    print(f'\n全量拟合自评(⚠️含泄漏, 非泛化): SC AUC={ate["auc"]:.3f} KS={ate["ks"]:.3f} '
          f'— test({split_year+1}+)已入训练, 勿当泛化能力')
    if loyo_stats:
        print(f'LOYO 去偏(✅真泛化): SC AUC={loyo_stats["auc_mean"]:.3f}±{loyo_stats["auc_std"]:.3f} | '
              f'KS={loyo_stats["ks_mean"]:.3f}±{loyo_stats["ks_std"]:.3f} ({loyo_stats["n_folds"]}折)')

    # 打印评分卡
    print_scorecard(kept, wbins, lr)

    # 入库
    cfg_tag = f'{_tag(horizon)}_{kind}_sc'
    ver = f'v_sc_{pd.Timestamp.now().strftime("%Y%m%d_%H%M")}_{cfg_tag}_{len(kept)}feat'
    # 训练集概率的 10 分位边界(9 切点): 部署后 predict 用它把新概率映射成固定 1-10 档
    # (训练标定的绝对档位, 跨批次可比; 10=训练集 top10% 概率)
    train_proba = lr.predict_proba(WOE_FILL(Xall_w))[:, 1]
    proba_deciles = np.quantile(train_proba, np.linspace(0.1, 0.9, 9)).tolist()
    print(f'  训练集概率 10 分位边界: {[round(d, 3) for d in proba_deciles]}')
    bundle = {'kind': 'scorecard', 'features': kept, 'woe_bins': wbins,
              'lr_model': lr, 'medians': {f: float(med[f]) for f in kept},
              'proba_deciles': proba_deciles}
    save_model_meta({
        'version': ver, 'label_config': cfg_tag, 'kind': kind, 'horizon': horizon, 'gray_cfg': gcfg,
        'features': kept, 'n_features': len(kept), 'medians': {f: float(med[f]) for f in kept},
        'lgb_model': None, 'lr_bundle': pickle.dumps(bundle),
        'metrics': {
            'sc_fit_auc': ate['auc'], 'sc_fit_ks': ate['ks'],   # 全量拟合自评(含泄漏, 非泛化)
            'sc_loyo_auc': loyo_stats['auc_mean'] if loyo_stats else None,
            'sc_loyo_auc_std': loyo_stats['auc_std'] if loyo_stats else None,
            'sc_loyo_ks': loyo_stats['ks_mean'] if loyo_stats else None,
            'sc_loyo_ks_std': loyo_stats['ks_std'] if loyo_stats else None,
            'sc_loyo_n_folds': loyo_stats['n_folds'] if loyo_stats else None,
        },
        'selection': sel,
        'selection_thresholds': {'iv_min': IV_MIN, 'psi_max': PSI_MAX, 'corr_max': CORR_MAX, 'vif_max': VIF_MAX},
        'n_train': int(len(yall)), 'dataset_version': 'derived_20260616_2334_f35ba6f3_7m',
        'note': f'WOE评分卡(标准五步特征) {cfg_tag}',
    })
    register_version('full', ver, ver,
                     metrics={'sc_fit_auc': ate['auc'], 'sc_fit_ks': ate['ks'],
                              'sc_loyo_auc': (loyo_stats or {}).get('auc_mean'),
                              'sc_loyo_ks': (loyo_stats or {}).get('ks_mean')},
                     n_features=len(kept), threshold=-10, n_samples=len(yall),
                     positive_rate=float(yall.mean()), files=['(in ml_model_meta DB)'],
                     note=f'WOE评分卡 {cfg_tag}', set_current=set_current,
                     label_config=cfg_tag, dataset_version='derived_20260616_2334_f35ba6f3_7m')
    cur = '✅ 已设为 current.full(生产)' if set_current else '(set_current=False, 未切生产)'
    print(f'\n入库: {ver} | {cur}')
    return ver


def main():
    ap = argparse.ArgumentParser(description='训练 WOE 评分卡并入库')
    ap.add_argument('features_path')
    ap.add_argument('--horizon', type=_parse_horizon, default=7)
    ap.add_argument('--kind', choices=['gray'], default='gray')
    ap.add_argument('--split-year', type=int, default=2024)
    ap.add_argument('--set-current', action='store_true', help='设为 current.full(生产)')
    ap.add_argument('--features', default=None, help='锁定特征(逗号分隔, 跳过select_features); 如共识特征')
    args = ap.parse_args()
    feats = args.features.split(',') if args.features else None
    run(args.features_path, args.horizon, args.kind, args.split_year, args.set_current, features=feats)


if __name__ == '__main__':
    main()
