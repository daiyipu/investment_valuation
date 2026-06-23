#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""任意期限 gray SC 模型的 LOYO per-year 报告(AUC/KS/IC + n), 对标 7m 报告。

每个版本出一份 output/validation_report_{tag}_gray_sc.md, 含:
  - 汇总(mean±std AUC/KS/IC, 特征数, 灰度阈值)
  - 逐年表: AUC(非灰) / KS(非灰) / IC(Spearman 概率 vs 收益, 全样本) / n

口径与 7m 报告一致:
  - AUC/KS: 当年【非灰】样本(赢>hi / 输<lo)的二分类区分度。
  - IC: 当年【全样本】(含灰)Spearman(预测概率, 连续收益) = 实战排序能力。
  - n: 当年非灰(有标签)样本数。
  - 每年留出、其余年训练 WOE+LR(同部署特征) → 真去偏 LOYO(非全量拟合)。

用法:
  python report_horizon.py --horizon 1                      # 最新 1m gray SC
  python report_horizon.py --horizon 1 --version v_sc_...   # 指定版本
"""
import argparse
import os
import sys
import pickle

import numpy as np
import pandas as pd
from scipy.stats import spearmanr
from sklearn.metrics import roc_auc_score
from sklearn.linear_model import LogisticRegression

HERE = os.path.dirname(os.path.abspath(__file__))                 # pipeline/(自身导入用)
ML_ROOT = os.path.dirname(HERE)                                    # ml_training/(data/output 在这)
sys.path.insert(0, HERE)
from db_model_store import load_predict_bundle, get_model_meta, list_model_metas
from eval_loyo import fit_woe, apply_woe
from validate_methods import make_features, calc_ks, eval_metrics
from train_horizon_models import build_label, GRAY_CFG, _prep, _parse_horizon

PARQUET = os.path.join(ML_ROOT, 'data', 'features_derived.parquet')
WOE_FILL = lambda X: X.replace([np.inf, -np.inf], np.nan).fillna(0)


def _ret_col(horizon):
    return f'{horizon[:-1]}周涨跌幅' if isinstance(horizon, str) and horizon.endswith('w') else f'{horizon}个月涨跌幅'


def _tag(horizon):
    return horizon if isinstance(horizon, str) else f'{horizon}m'


def latest_gray_sc(horizon):
    """该期限最新 gray SC 版本(无 lgb_model = 纯评分卡)。"""
    metas = [m for m in list_model_metas(kind='gray')
             if not m.get('lgb_model') and m.get('kind') == 'gray'
             and str(m.get('horizon')) == str(horizon)]
    if not metas:
        return None
    metas.sort(key=lambda m: m.get('version', ''))
    return metas[-1]['version']


def per_year_loyo(df, horizon, features, lo, hi):
    """逐年 LOYO SC: 每年留出, 其余年训练 WOE+LR(锁定特征), 报 AUC/KS/IC。

    训练用非灰样本(make_features 自动丢灰度 NaN 标签), 打分/IC 用当年【全行】(含灰),
    与 7m 报告口径一致: AUC/KS=非灰, IC=全样本 Spearman(概率, 收益)。"""
    ret_col = _ret_col(horizon)
    df = df.dropna(subset=['报价日']).reset_index(drop=True)
    df['_y'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000).astype('Int64')
    years = sorted(int(y) for y in df['_y'].dropna().unique())
    rows = []
    for Y in years:
        dtr = df[df['_y'] != Y].drop(columns=['_y'])
        dte = df[df['_y'] == Y].drop(columns=['_y'])
        lbl, _ = build_label(dtr, horizon, 'gray')
        # 训练: 非灰样本(make_features 丢灰度)
        Xtr_raw, ytr, _ = make_features(dtr, label_col=lbl, ret_col=ret_col)
        if Xtr_raw is None or ytr.nunique() < 2:
            rows.append({'year': Y, 'auc': np.nan, 'ks': np.nan, 'ic': np.nan, 'n': 0})
            continue
        Xtr, med = _prep(Xtr_raw)
        feats = [f for f in features if f in Xtr.columns]
        Xtr_w, wb = fit_woe(Xtr[feats], ytr, feats)
        lr = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)
        lr.fit(WOE_FILL(Xtr_w), ytr)
        # 打分: 当年全行(含灰) → median 填 → apply_woe → LR
        Xte = pd.DataFrame(index=dte.index)
        for f in feats:
            Xte[f] = pd.to_numeric(dte[f], errors='coerce') if f in dte.columns else np.nan
        Xte = Xte.fillna({f: med.get(f, 0) for f in feats}).replace([np.inf, -np.inf], 0)
        Xte_w = apply_woe(Xte, feats, wb)
        p = lr.predict_proba(WOE_FILL(Xte_w))[:, 1]
        r = pd.to_numeric(dte[ret_col], errors='coerce')
        # 非灰 AUC/KS(eval_metrics 统一口径: n<20 或单类 → nan 自动丢, 与 loyo_fixed 一致)
        ng = (r > hi) | (r < lo)
        y_ng = (r[ng] > hi).astype(int).values
        _em = eval_metrics(y_ng, p[ng]) if ng.sum() else None
        auc = _em['auc'] if _em else np.nan
        ks = _em['ks'] if _em else np.nan
        # 全样本 IC(含灰)
        v = np.isfinite(r.values) & np.isfinite(p)
        ic = float(spearmanr(p[v], r.values[v]).correlation) if v.sum() > 20 else np.nan
        rows.append({'year': Y, 'auc': auc, 'ks': ks, 'ic': ic, 'n': int(ng.sum())})
    return rows


def tier_distribution(df, horizon, ver, lo, hi):
    """全样本按 档位(训练集概率10分位)分桶, 报每档实际胜率/收益分布(in-sample 校准视图)。
    档位 = 部署模型的 proba_deciles 把预测概率切成 10 档(10=训练集 top10% 概率)。"""
    b = pickle.loads(load_predict_bundle(ver)['lr_bundle'])
    feats, medians, lr = b['features'], b.get('medians', {}), b['lr_model']
    dec = b.get('proba_deciles', [])
    X = pd.DataFrame(index=df.index)
    for f in feats:
        X[f] = pd.to_numeric(df[f], errors='coerce') if f in df.columns else np.nan
    X = X.fillna({f: medians.get(f, 0) for f in feats}).replace([np.inf, -np.inf], 0)
    Xw = apply_woe(X, feats, b['woe_bins']).replace([np.inf, -np.inf], 0).fillna(0)
    p = lr.predict_proba(Xw)[:, 1]
    tier = (np.digitize(p, dec) + 1) if dec else np.ones(len(p), dtype=int)
    ret = pd.to_numeric(df[_ret_col(horizon)], errors='coerce')
    t = pd.DataFrame({'tier': tier, 'p': p, 'ret': ret}).dropna(subset=['ret'])
    rows = []
    for k in range(10, 0, -1):
        g = t[t['tier'] == k]
        if len(g) == 0:
            continue
        plo = dec[k - 2] if k >= 2 else 0.0
        phi = dec[k - 1] if k <= 9 else 1.0
        rows.append({'tier': k, 'n': len(g), 'plo': plo, 'phi': phi,
                     'win': float((g['ret'] > hi).mean()),
                     'lose': float((g['ret'] < lo).mean()),
                     'gray': float(((g['ret'] >= lo) & (g['ret'] <= hi)).mean()),
                     'mean': float(g['ret'].mean()), 'median': float(g['ret'].median())})
    return rows, len(t), float(t['ret'].mean()), float(t['ret'].median())


def generate_report(features_path, horizon, version=None):
    """生成某期限(最新或指定版本)gray SC 的 per-year 报告 md。
    供 train_to_production 入库后自动调用(管线第 ⑩步后), 也可 CLI 单跑。返回报告路径。"""
    lo, hi = GRAY_CFG[horizon]
    ver = version or latest_gray_sc(horizon)
    if not ver:
        print(f'❌ 找不到 {horizon}m gray SC 模型'); return None
    meta = get_model_meta(ver) or {}
    feats = meta.get('features') or []
    if not feats:
        feats = pickle.loads(load_predict_bundle(ver)['lr_bundle'])['features']
    print(f'\n=== 生成报告: {ver} | {_tag(horizon)} gray({lo}/{hi}) | {len(feats)}特征 ===')
    df = pd.read_parquet(features_path)
    rows = per_year_loyo(df, horizon, feats, lo, hi)
    aucs = [r['auc'] for r in rows if r['auc'] == r['auc']]
    kss = [r['ks'] for r in rows if r['ks'] == r['ks']]
    ics = [r['ic'] for r in rows if r['ic'] == r['ic']]
    am, asd = float(np.mean(aucs)), float(np.std(aucs))
    km, ksd = float(np.mean(kss)), float(np.std(kss))
    im, isd = float(np.mean(ics)), float(np.std(ics))
    n_tot = sum(r['n'] for r in rows)
    print(f'汇总: AUC {am:.3f}±{asd:.3f} | KS {km:.3f}±{ksd:.3f} | IC {im:+.4f}±{isd:.4f} | 非灰样本 {n_tot}')
    print(f'{"年":<6}{"AUC":>8}{"KS":>8}{"IC":>10}{"n":>7}')
    for r in rows:
        ic_s = f"{r['ic']:+.4f}" if r['ic'] == r['ic'] else '   nan'
        au_s = f"{r['auc']:.3f}" if r['auc'] == r['auc'] else '  nan'
        ks_s = f"{r['ks']:.3f}" if r['ks'] == r['ks'] else '  nan'
        print(f"{r['year']:<6}{au_s:>8}{ks_s:>8}{ic_s:>10}{r['n']:>7}")

    trows, tn, tmean, tmed = tier_distribution(df, horizon, ver, lo, hi)
    _ret = pd.to_numeric(df[_ret_col(horizon)], errors='coerce')
    bw, bl = float((_ret > hi).mean()), float((_ret < lo).mean())
    print(f'\n档位盈利分布(in-sample n={tn}, 基线 胜>{hi}%={bw*100:.1f}% / 输<{lo}%={bl*100:.1f}% / 均{tmean:+.2f}%):')
    print(f'{"档":>3}{"n":>6}{"概率区间":>16}{"胜率":>8}{"输率":>8}{"均收益":>9}{"中位":>8}')
    for r in trows:
        print(f'{r["tier"]:>3}{r["n"]:>6}  [{r["plo"]:.3f},{r["phi"]:.3f})'
              f'{r["win"]*100:>7.1f}%{r["lose"]*100:>7.1f}%{r["mean"]:>+8.2f}%{r["median"]:>+7.2f}%')

    out = os.path.join(ML_ROOT, 'output', f'validation_report_{_tag(horizon)}_gray_sc.md')
    derived = meta.get('dataset_version', '?')
    n_train = meta.get('n_train', '?')
    L = [f'# {_tag(horizon)} gray SC 验证报告(LOYO per-year AUC/KS/IC)\n',
         f'版本 {ver} | {len(feats)}特征 | 灰度阈值(输/赢) {lo}/{hi} | 训练样本 {n_train} | 数据 {derived}\n',
         f'生成 {pd.Timestamp.now().strftime("%Y-%m-%d")}\n',
         f'\nIC = 预测概率 vs 实际{_tag(horizon)}涨跌幅 Spearman(当年全样本, 含灰)\n',
         '\n## 汇总\n', '| 指标 | mean±std |', '|---|---|',
         f'| AUC | {am:.3f}±{asd:.3f} |', f'| KS | {km:.3f}±{ksd:.3f} |',
         f'| IC | {im:+.4f}±{isd:.4f} |', '\n## 逐年\n', '| 年 | AUC | KS | IC | n |',
         '|---|---|---|---|---|']
    for r in rows:
        ic_s = f"{r['ic']:+.4f}" if r['ic'] == r['ic'] else 'nan'
        au_s = f"{r['auc']:.3f}" if r['auc'] == r['auc'] else 'nan'
        ks_s = f"{r['ks']:.3f}" if r['ks'] == r['ks'] else 'nan'
        L.append(f"| {r['year']} | {au_s} | {ks_s} | {ic_s} | {r['n']} |")
    L += [f'\n## 档位盈利概率分布(全样本 in-sample, 按实际{_tag(horizon)}收益)\n',
          f'基线: n={tn} | 胜率(>{hi}%) {bw*100:.1f}% | 输率(<{lo}%) {bl*100:.1f}% | 均收益 {tmean:+.2f}%\n',
          '档位 = 训练集预测概率 10 分位(10=top10% 概率)。校准判据: 高档位应胜率更高、输率更低。\n',
          '| 档 | n | 概率区间 | 胜率% | 输率% | 灰区% | 均收益% | 中位% |',
          '|---|---|---|---|---|---|---|---|']
    for r in trows:
        L.append(f"| {r['tier']} | {r['n']} | [{r['plo']:.3f},{r['phi']:.3f}) | "
                 f"{r['win']*100:.1f} | {r['lose']*100:.1f} | {r['gray']*100:.1f} | "
                 f"{r['mean']:+.2f} | {r['median']:+.2f} |")
    L += [f'\n## 入模特征({len(feats)})\n', '`' + '`, `'.join(feats) + '`\n']
    with open(out, 'w', encoding='utf-8') as f:
        f.write('\n'.join(L))
    print(f'写出: {out}')
    return out


def main():
    ap = argparse.ArgumentParser(description='gray SC 期限模型 LOYO per-year 报告')
    ap.add_argument('features_path', nargs='?', default=PARQUET, help='features_derived.parquet')
    ap.add_argument('--horizon', type=_parse_horizon, default=1)
    ap.add_argument('--version', default=None, help='指定版本(默认该期限最新)')
    args = ap.parse_args()
    generate_report(args.features_path, args.horizon, args.version)


if __name__ == '__main__':
    main()
