#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""LGB vs SC × 共识字段 vs 全量字段 对比(薄编排, 复用管线函数, 不重写训练/评估)。

- 共识字段 = 该期限最新 gray SC 的入模特征(train_to_production 的 共识→LOO→final 产出)。
- 全量字段 = make_features 全部特征池(仅去泄漏/标签/收益, 无 IV/共识精简)。
- 对两组特征各跑 loyo_fixed → 同一 LOYO 口径下的 LGB/LR/SC AUC/KS, 直接对比。
- --deploy-lgb: 用共识字段调 deploy_lgb 部署一个共识-LGB(set_current=False, 不碰 SC 生产)。

充分复用: loyo_fixed(eval_loyo) + deploy_lgb(train_to_production) + make_features(validate_methods)。
不重新实现任何训练/评估逻辑。

用法:
  python ml_training/compare_lgb_sc.py --horizon 7 --deploy-lgb
"""
import argparse
import contextlib
import io
import os
import sys

import pandas as pd

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # ml_training/(搬自 report/ 升一层)            # ml_training/
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, 'pipeline'))           # 管线模块
from validate_methods import make_features
from train.train_horizon_models import build_label, _ret_col, _tag, _parse_horizon
from eval_loyo import loyo_fixed
from train.train_to_production import deploy_lgb
from deploy.db_model_store import get_model_meta
from report.report_horizon import latest_gray_sc


def full_features(df, horizon):
    """make_features 全部特征列(去泄漏/标签/收益), 再剔 全NaN/常数/>50%缺失列
    (这些列在某些折 train 上全 NaN → median NaN → LR/StandardScaler 崩;
     共识特征经选择无此问题, 仅全量池需预清洗)。"""
    lbl, _ = build_label(df, horizon, 'gray')
    X, _, _ = make_features(df, label_col=lbl, ret_col=_ret_col(horizon))
    miss = X.isna().mean()
    nunique = X.nunique()
    return [c for c in X.columns if miss[c] < 0.5 and nunique[c] > 1]


def run_compare(features_path, horizon, deploy=False):
    df = pd.read_parquet(features_path)
    ver = latest_gray_sc(horizon)
    cons = get_model_meta(ver)['features']
    full = full_features(df, horizon)
    print(f'{_tag(horizon)} gray | 共识 {len(cons)}特征({ver}) vs 全量 {len(full)}特征\n')

    res = {}
    for name, feats in [('共识', cons), ('全量', full)]:
        with contextlib.redirect_stdout(io.StringIO()):
            res[name] = loyo_fixed(features_path, horizon, 'gray', ','.join(feats))

    print('=' * 98)
    print(f'{_tag(horizon)} 对比: LGB / LR / SC  ×  共识({len(cons)}) vs 全量({len(full)})  (LOYO; 非灰/含灰双标签)')
    print('=' * 98)
    print(f'{"模型":<5}{"标签":<5}{"共识AUC±std":>14}{"KS":>7}{"全量AUC±std":>14}{"KS":>7}{"ΔAUC":>8}{"IC共识":>9}{"IC全量":>9}')
    print('-' * 98)
    for m, lab in (('lgb', 'LGB'), ('lr', 'LR'), ('sc', 'SC')):
        c, f = res['共识'][m], res['全量'][m]
        for tag, ak, kk in (('非灰', 'auc_mean', 'ks_mean'), ('含灰', 'incl_auc_mean', 'incl_ks_mean')):
            cv, fv = c.get(ak), f.get(ak)
            da = (cv - fv) if cv is not None and fv is not None else float('nan')
            cs = f'{cv:.3f}±{c.get(ak.replace("mean","std"),0):.3f}' if cv is not None else '   nan   '
            fs = f'{fv:.3f}±{f.get(ak.replace("mean","std"),0):.3f}' if fv is not None else '   nan   '
            print(f'{lab:<5}{tag:<5}{cs:>20}{(c.get(kk) or 0):>7.3f}{fs:>18}{(f.get(kk) or 0):>7.3f}'
                  f'{da:>+8.3f}{(c.get("ic_mean") or 0):>+9.3f}{(f.get("ic_mean") or 0):>+9.3f}')
        print()
    print('读法: ΔAUC>0=共识优于全量; 非灰=已决样本区分力, 含灰=全样本实战口径(阈值7m=-10%/其余0)。'
          'IC=全样本 Spearman(两标签共用)。')

    if deploy:
        print(f'\n--- 用共识字段部署 LGB(deploy_lgb, set_current=False) ---')
        v = deploy_lgb(features_path, horizon, 'gray', cons, 2024, False)
        print(f'部署: {v} | 与共识-SC({ver})同 15 特征, 入库不切生产。')


def main():
    ap = argparse.ArgumentParser(description='LGB/SC × 共识/全量 对比(复用管线)')
    ap.add_argument('--horizon', type=_parse_horizon, default=7)
    ap.add_argument('--features-path',
                    default=os.path.join(HERE, 'data', 'features_derived.parquet'))
    ap.add_argument('--deploy-lgb', action='store_true', help='用共识字段部署 LGB(不切生产)')
    args = ap.parse_args()
    run_compare(args.features_path, args.horizon, args.deploy_lgb)


if __name__ == '__main__':
    main()
