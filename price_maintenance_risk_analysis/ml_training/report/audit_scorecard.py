#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""评分卡审计报告 —— 上线前人工核查基础。

为何: 评分卡上线前审核人员需看到每个指标的 业务解释/IV/贡献度/分箱/分值,
      不能只看一个 AUC。本脚本把 current.full(或指定版本)的 SC 评分卡
      全部明细导出存档(CSV + Markdown)。

每特征输出:
  业务解释 | IV | LR系数 | 贡献度(logit摆幅=coef×(maxWOE-minWOE)) | 贡献占比% |
  分箱边界 | 各箱WOE | 各箱分值(=coef×WOE×100)

数据来源: lr_bundle(woe_bins/lr_model/proba_deciles) + features_derived 重算 IV。
业务解释: feature_glossary.explain()。

用法:
  python ml_training/audit_scorecard.py                    # 审计 current.full
  python ml_training/audit_scorecard.py <version>
"""
import os
import sys
import pickle
import argparse

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # ml_training/
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'pipeline'))   # 管线模块已移入 pipeline/
from model_registry import get_current
from db_model_store import load_predict_bundle, get_model_meta
from train_scorecard import calc_iv_all_features
from report.feature_glossary import explain

PARQUET = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'features_derived.parquet')


def _gray_label(df, horizon=7):
    """horizon 个月 gray 标签, 用生产口径 GRAY_CFG(与训练一致)。"""
    from train_horizon_models import GRAY_CFG
    lo, hi = GRAY_CFG.get(horizon, (-20, 10))
    col = f'{horizon}个月涨跌幅'
    s = pd.to_numeric(df[col], errors='coerce') if col in df.columns else pd.Series(np.nan, index=df.index)
    return pd.Series(np.where(s > hi, 1, np.where(s < lo, 0, np.nan)), index=df.index)


def audit(version, out_dir):
    bundle = load_predict_bundle(version)
    sc = pickle.loads(bundle['lr_bundle'])
    feats = sc['features']
    woe_bins = sc['woe_bins']
    lr = sc['lr_model']
    coefs = dict(zip(feats, lr.coef_[0]))
    intercept = float(lr.intercept_[0])
    meta = get_model_meta(version) or {}
    horizon = meta.get('horizon', 7)

    # 重算 IV(训练集, 同期限灰度标签)
    df = pd.read_parquet(PARQUET)
    y = _gray_label(df, horizon)
    X = df[feats].apply(pd.to_numeric, errors='coerce')
    iv_tbl = calc_iv_all_features(X, y)
    iv_map = dict(zip(iv_tbl['feature'], iv_tbl['iv']))

    rows = []
    for f in feats:
        wb = woe_bins.get(f, {})
        rights = list(wb.get('rights', []))
        woes = [float(w) for w in wb.get('woes', [])]
        coef = float(coefs.get(f, 0.0))
        pts = [round(coef * w * 100, 2) for w in woes]          # 各箱分值 = coef×WOE×100
        swing = float(coef * (max(woes) - min(woes))) if woes else 0.0
        rows.append({
            '特征': f, '业务解释': explain(f),
            'IV': round(float(iv_map.get(f, np.nan)), 4),
            'LR系数': round(coef, 4),
            '贡献度_logit摆幅': round(swing, 4),
            '分箱边界': [round(r, 4) for r in rights],
            '各箱WOE': [round(w, 3) for w in woes],
            '各箱分值_x100': pts,
        })
    adf = pd.DataFrame(rows)
    tot = adf['贡献度_logit摆幅'].abs().sum() or 1.0
    adf['贡献度占比%'] = (adf['贡献度_logit摆幅'].abs() / tot * 100).round(1)
    adf = adf.reindex(adf['贡献度_logit摆幅'].abs().sort_values(ascending=False).index)

    os.makedirs(out_dir, exist_ok=True)
    csv_p = os.path.join(out_dir, 'scorecard_audit.csv')
    adf.to_csv(csv_p, index=False, encoding='utf-8-sig')

    dec = [round(float(x), 3) for x in sc.get('proba_deciles', [])]
    md = [f'# 评分卡审计报告 — {version}', '',
          f'- 期限: {horizon}m gray | 特征数: {len(feats)} | 截距(intercept): {intercept:+.4f}',
          f'- 模型 LOYO: AUC {meta.get("sc_loyo_auc")} | KS {meta.get("sc_loyo_ks")}'
          f'  (若空则仅有 OOT: AUC {meta.get("sc_oot_auc")} / KS {meta.get("sc_oot_ks")})',
          f'- 档位边界(proba 10 分位): {dec}', '',
          '> 分值 = LR系数 × 该箱WOE × 100(logit 贡献量, 越大越推高盈利概率);',
          '> 贡献度 = 系数×(最高箱WOE−最低箱WOE), 即该特征能造成的 logit 最大摆幅;',
          '> 贡献占比 = 该特征|摆幅|占全部特征之和的比例。', '',
          '| 特征 | 业务解释 | IV | LR系数 | 贡献度(摆幅) | 占比% | 分箱边界 | 各箱WOE | 各箱分值×100 |',
          '|---|---|---|---|---|---|---|---|---|']
    for _, r in adf.iterrows():
        md.append(f'| {r["特征"]} | {r["业务解释"]} | {r["IV"]} | {r["LR系数"]:+.4f} | '
                  f'{r["贡献度_logit摆幅"]:+.4f} | {r["贡献度占比%"]:.1f} | '
                  f'{r["分箱边界"]} | {r["各箱WOE"]} | {r["各箱分值_x100"]} |')
    md_p = os.path.join(out_dir, 'scorecard_audit.md')
    with open(md_p, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(md))
    print(f'✅ 审计报告已存档: {out_dir}/  (scorecard_audit.csv + scorecard_audit.md)\n')
    show = adf[['特征', '业务解释', 'IV', 'LR系数', '贡献度_logit摆幅', '贡献度占比%']].copy()
    print(show.to_string(index=False))
    return adf


def main():
    ap = argparse.ArgumentParser(description='评分卡审计报告(业务解释+IV+贡献度+分箱+分值)')
    ap.add_argument('version', nargs='?', help='模型版本(默认 current.full)')
    ap.add_argument('--out', help='输出目录(默认 output/audit_<version>)')
    args = ap.parse_args()
    version = args.version or get_current('full')
    out_dir = args.out or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       'output', f'audit_{version}')
    audit(version, out_dir)


if __name__ == '__main__':
    main()
