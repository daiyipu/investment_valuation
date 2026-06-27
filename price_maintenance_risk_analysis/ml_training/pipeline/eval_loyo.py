#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""【标准评估】Leave-One-Year-Out (LOYO) —— 项目唯一评估口径(2026-06-16 起替代单 OOT)。

单 OOT(split≤Y/test>Y) 只看一个 regime 的测试年, 数值偏且小差异在噪声内 → 已弃用
(sweep_label_horizons.py 已删)。LOYO 挖掉某一年作 test、其余年训练, 6 年轮流 →
跨 regime 的 mean±std 才是去偏的可信估计。选配置/比模型/定主力一律看本脚本输出。
train_horizon_models 内的单 OOT 仅作 sanity, 不作取舍依据。

解决两个问题:
  1. 样本量: 每折用 5 年(~1364 样本)训练, 远多于单 split 的尾巴。
  2. 单 OOT 偏差: 6 年(熊/震荡/牛)轮流当测试 → 跨 regime mean±std, 不被某一年误导。

三种模型同特征集对比(选择统一走标准五步 select_features):
  LGB = LightGBM 树(原始值, 树量纲无关)
  LR  = 逻辑回归(StandardScaler 标准化原始值)
  SC  = 评分卡(WOE 分箱变换 + 逻辑回归; 抗长尾/可解释成评分点)

⚠ LOYO 训练集含测试年的相邻年 → CV 式泛化估计, 略乐观, 非严格前向回测。
   WOE 用自包含 fit_woe/apply_woe: train 拟合分箱存边+woe, test 用 searchsorted 套 train 边
   (不重算 WOE, 不碰 test 标签, 防泄漏)。

用法: python eval_loyo.py <features_derived.parquet>
输出: output/loyo_{lgb,lr,sc}.csv (per-year AUC/KS + mean±std)
"""
import argparse
import os
import sys

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_methods import make_features, eval_metrics
from features.feature_selection import select_features, prune_by_lgb_importance
from train.train_horizon_models import GRAY_CFG, HORIZONS, build_label, _prep, _train, _ret_col, _tag, _parse_horizon


# ─────────────── 自包含 WOE(防 test 标签泄漏) ───────────────
def fit_woe(Xtr, ytr, feats, n_bins=5):
    """train 拟合: qcut 分箱 → 每 bin 算 WOE → 存 rights(升序)+woes。返回 (Xtr_woe, bins)。"""
    bins, Xw = {}, Xtr.copy()
    for f in feats:
        s, yv = Xtr[f], ytr
        valid = s.notna() & yv.notna()
        sv = s[valid]
        if len(sv) < 50:
            continue
        try:
            binned = pd.qcut(sv, n_bins, duplicates='drop')
        except ValueError:
            continue
        d = pd.DataFrame({'b': binned, 'y': yv[valid].values})
        tp, tn = max(d.y.sum(), 1), max(len(d) - d.y.sum(), 1)
        edges, woes = [], []
        for name, g in d.groupby('b', observed=True):
            pos, neg = g.y.sum(), len(g) - g.y.sum()
            woes.append(np.log(max(pos / tp, 1e-4) / max(neg / tn, 1e-4)))
            edges.append((float(name.left), float(name.right)))
        if not edges:
            continue
        order = np.argsort([e[0] for e in edges])
        rights = [edges[i][1] for i in order]
        woes = [woes[i] for i in order]
        bins[f] = {'rights': rights, 'woes': woes}
        Xw[f] = _apply_bin(Xtr[f].values, rights, woes)
    return Xw, bins


def _apply_bin(vals, rights, woes):
    rights, woes = np.asarray(rights), np.asarray(woes)
    idx = np.clip(np.searchsorted(rights, vals, side='left'), 0, len(woes) - 1)
    out = np.where(pd.notna(vals), woes[idx], np.nan)
    return out


def apply_woe(X, feats, bins):
    """用 train 的 bins 把新数据 X 变换成 WOE(不碰 y)。"""
    Xw = X.copy()
    for f in feats:
        if f in bins and f in X.columns:
            Xw[f] = _apply_bin(X[f].values, bins[f]['rights'], bins[f]['woes'])
        elif f in X.columns:
            Xw[f] = np.nan
    return Xw


# ─────────────── 单折三模型评估 ───────────────
def _eval_fold(dtr, dte, horizon, kind):
    """返回 dict(lgb_auc/ks, lr_auc/ks, sc_auc/ks, n_feat) 或 None。"""
    lbl, _ = build_label(dtr, horizon, kind)
    if kind == 'gray':
        build_label(dte, horizon, kind)
    ret = _ret_col(horizon)
    Xtr_raw, ytr, _ = make_features(dtr, label_col=lbl, ret_col=ret)
    Xte_raw, yte, _ = make_features(dte, label_col=lbl, ret_col=ret)
    if Xtr_raw is None or Xte_raw is None or len(ytr) < 40 or yte.nunique() < 2:
        return None
    Xtr, med = _prep(Xtr_raw)
    Xte = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr.columns)

    # 标准五步选择(三模型共用)
    kept, _ = select_features(Xtr, ytr, Xte, Xtr_raw=Xtr_raw, Xte_raw=Xte_raw)
    gbm, lr, sc_lr = _train(Xtr[kept], ytr)          # LGB + LR(StandardScaler)
    final = prune_by_lgb_importance(gbm, kept)
    if 5 <= len(final) < len(kept):
        kept = final
        gbm, lr, sc_lr = _train(Xtr[kept], ytr)
    Xtr_s, Xte_s = Xtr[kept], Xte[kept]

    al = eval_metrics(yte.values, gbm.predict_proba(Xte_s)[:, 1])
    ar = eval_metrics(yte.values, lr.predict_proba(sc_lr.transform(Xte_s))[:, 1])

    # 评分卡 SC: WOE 变换 + LR(同 kept 特征)
    Xtr_woe, wbins = fit_woe(Xtr_s, ytr, kept)
    Xte_woe = apply_woe(Xte_s, kept, wbins)
    from sklearn.linear_model import LogisticRegression
    sc_model = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)
    sc_model.fit(Xtr_woe.replace([np.inf, -np.inf], np.nan).fillna(0), ytr)
    asc = eval_metrics(yte.values, sc_model.predict_proba(
        Xte_woe.replace([np.inf, -np.inf], np.nan).fillna(0))[:, 1])

    return {'lgb_auc': al['auc'], 'lgb_ks': al['ks'],
            'lr_auc': ar['auc'], 'lr_ks': ar['ks'],
            'sc_auc': asc['auc'], 'sc_ks': asc['ks'], 'n_feat': len(kept)}


def run(features_path):
    df = pd.read_parquet(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    df['_year'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000).astype('Int64')
    years = sorted(int(y) for y in df['_year'].dropna().unique())
    print(f'LOYO {len(years)} 折: {years} | 全量 {len(df)} 行 | 三模型(LGB/LR/SC)同标准五步特征\n')

    configs = [(f'{h}m_{k}', h, k) for h in HORIZONS for k in ('gray',)]
    models = ('lgb', 'lr', 'sc')
    # res[model][metric] = list of rows(cfg, gray, {year:val}, mean, std)
    res = {m: {'auc': [], 'ks': []} for m in models}

    for name, h, kind in configs:
        gtag = f'{GRAY_CFG[h][0]}/{GRAY_CFG[h][1]}' if kind == 'gray' else '-'
        auc_row = {m: {'cfg': name, 'gray': gtag} for m in models}
        ks_row = {m: {'cfg': name, 'gray': gtag} for m in models}
        for Y in years:
            dtr = df[df['_year'] != Y].drop(columns=['_year'])
            dte = df[df['_year'] == Y].drop(columns=['_year'])
            r = _eval_fold(dtr, dte, h, kind)
            for m in models:
                auc_row[m][Y] = round(r[f'{m}_auc'], 3) if r else np.nan
                ks_row[m][Y] = round(r[f'{m}_ks'], 3) if r else np.nan
        for m in models:
            for row, key in ((auc_row[m], 'auc'), (ks_row[m], 'ks')):
                vals = [row[Y] for Y in years if pd.notna(row.get(Y))]
                row['mean'] = round(np.mean(vals), 3) if vals else np.nan
                row['std'] = round(np.std(vals), 3) if vals else np.nan
                res[m][key].append(row)
        print(f"  {name:<9} gray={gtag:<6} | AUC  LGB {auc_row['lgb']['mean']}±{auc_row['lgb']['std']}  "
              f"LR {auc_row['lr']['mean']}±{auc_row['lr']['std']}  SC {auc_row['sc']['mean']}±{auc_row['sc']['std']}")

    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'output')
    for m in models:
        pd.DataFrame(res[m]['auc']).to_csv(os.path.join(out_dir, f'loyo_{m}.csv'), index=False)
        pd.DataFrame(res[m]['ks']).to_csv(os.path.join(out_dir, f'loyo_{m}_ks.csv'), index=False)

    print('\n' + '=' * 98)
    for m, label in (('lgb', 'LGB'), ('lr', 'LR'), ('sc', 'SC 评分卡')):
        d = pd.DataFrame(res[m]['auc']).dropna(subset=['mean']).sort_values('mean', ascending=False)
        print(f'\n【{label}】LOYO AUC mean±std 排名 (per-year 见 loyo_{m}.csv):')
        print(d[['cfg', 'gray', *years, 'mean', 'std']].to_string(index=False))

    print('\n【每配置最强模型】(按 LOYO mean AUC):')
    by_cfg = {}
    for m in models:
        for r in res[m]['auc']:
            by_cfg.setdefault(r['cfg'], {})[m] = r['mean']
    for cfg, ms in by_cfg.items():
        best = max(ms, key=ms.get)
        print(f"  {cfg:<9} → {best.upper()} {ms[best]}  (LGB {ms['lgb']} / LR {ms['lr']} / SC {ms['sc']})")
    print(f'\n写出: {out_dir}/loyo_{{lgb,lr,sc}}.csv (+_ks.csv)')


def _ic_full(dte, feats, med, gbm, lr, sc_lr, sc_m, wb, ret_col, incl_thr):
    """全 dte 样本(含灰)上, per LGB/LR/SC: IC(Spearman) + 含灰标签 AUC/KS(单阈值二分)。
    在全行打分(median 填 → apply_woe), 不限非灰; IC/含灰都用全行预测, 对齐 ret。"""
    from scipy.stats import spearmanr
    r = pd.to_numeric(dte[ret_col], errors='coerce') if ret_col in dte.columns else pd.Series(np.nan, index=dte.index)
    Xf = pd.DataFrame(index=dte.index)
    for f in feats:
        Xf[f] = pd.to_numeric(dte[f], errors='coerce') if f in dte.columns else np.nan
    Xf = Xf.fillna({f: med.get(f, 0) for f in feats}).replace([np.inf, -np.inf], 0)
    v = np.isfinite(r.values)
    out = {m: {'ic': np.nan, 'incl_auc': np.nan, 'incl_ks': np.nan} for m in ('lgb', 'lr', 'sc')}
    if v.sum() > 20:
        rv = r.values[v]
        Xfw = apply_woe(Xf, feats, wb).replace([np.inf, -np.inf], 0).fillna(0)
        preds = {'lgb': gbm.predict_proba(Xf)[:, 1],
                 'lr': lr.predict_proba(sc_lr.transform(Xf))[:, 1],
                 'sc': sc_m.predict_proba(Xfw)[:, 1]}
        y_incl = (rv > incl_thr).astype(int)
        for name, p in preds.items():          # p 全行, [v] 取有限 ret 子集(与 rv/y_incl 对齐)
            out[name]['ic'] = float(spearmanr(p[v], rv).correlation)
            _em = eval_metrics(y_incl, p[v])
            out[name]['incl_auc'] = _em['auc']
            out[name]['incl_ks'] = _em['ks']
    return out


def _eval_fold_fixed(dtr, dte, horizon, kind, features):
    """锁定特征的单折评估(跳过 select_features), 返回 lgb/lr/sc auc/ks/ic。
    用于部署模型的"对齐验证"——部署用固定特征集, 验证也用同一套, 避免 选/部 偏离。"""
    lbl, _ = build_label(dtr, horizon, kind)
    if kind == 'gray':
        build_label(dte, horizon, kind)
    ret = _ret_col(horizon)
    Xtr_raw, ytr, _ = make_features(dtr, label_col=lbl, ret_col=ret)
    Xte_raw, yte, _ = make_features(dte, label_col=lbl, ret_col=ret)
    if Xtr_raw is None or Xte_raw is None or yte.nunique() < 2:
        return None
    Xtr, med = _prep(Xtr_raw)
    Xte = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr.columns)
    feats = [f for f in features if f in Xtr.columns]
    Xtr_s, Xte_s = Xtr[feats], Xte[feats]
    gbm, lr, sc_lr = _train(Xtr_s, ytr)
    al = eval_metrics(yte.values, gbm.predict_proba(Xte_s)[:, 1])
    ar = eval_metrics(yte.values, lr.predict_proba(sc_lr.transform(Xte_s))[:, 1])
    Xtr_w, wb = fit_woe(Xtr_s, ytr, feats)
    Xte_w = apply_woe(Xte_s, feats, wb)
    from sklearn.linear_model import LogisticRegression
    m = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)
    m.fit(Xtr_w.replace([np.inf, -np.inf], np.nan).fillna(0), ytr)
    asc = eval_metrics(yte.values, m.predict_proba(
        Xte_w.replace([np.inf, -np.inf], np.nan).fillna(0))[:, 1])
    from train.train_horizon_models import _incl_thr
    full = _ic_full(dte, feats, med, gbm, lr, sc_lr, m, wb, ret, _incl_thr(horizon))  # 全行 IC + 含灰 AUC/KS
    return {'lgb_auc': al['auc'], 'lgb_ks': al['ks'], 'lr_auc': ar['auc'],
            'lr_ks': ar['ks'], 'sc_auc': asc['auc'], 'sc_ks': asc['ks'], 'n_feat': len(feats),
            'lgb_ic': full['lgb']['ic'], 'lr_ic': full['lr']['ic'], 'sc_ic': full['sc']['ic'],
            'lgb_auc_incl': full['lgb']['incl_auc'], 'lgb_ks_incl': full['lgb']['incl_ks'],
            'lr_auc_incl': full['lr']['incl_auc'], 'lr_ks_incl': full['lr']['incl_ks'],
            'sc_auc_incl': full['sc']['incl_auc'], 'sc_ks_incl': full['sc']['incl_ks']}


def loyo_fixed(features_path, horizon, kind, features):
    """锁定特征 LOYO(部署对齐验证): 6 折每折用同一套固定特征训 LGB/LR/SC, 报 mean±std。
    用法: python eval_loyo.py <parquet> --fixed-features "f1,f2,..." --horizon 7 --kind gray"""
    df = pd.read_parquet(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    df['_year'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000).astype('Int64')
    years = sorted(int(y) for y in df['_year'].dropna().unique())
    feats = [f.strip() for f in features.split(',')]
    print(f'🔒 锁定特征 LOYO [{len(feats)}特征]: {feats}\n  期限 {_tag(horizon)}/{kind} | {years}\n')
    models = ('lgb', 'lr', 'sc')
    per = {m: [] for m in models}
    per_ks = {m: [] for m in models}
    per_ic = {m: [] for m in models}
    per_ai = {m: [] for m in models}   # 含灰 AUC
    per_ki = {m: [] for m in models}   # 含灰 KS
    for Y in years:
        dtr = df[df['_year'] != Y].drop(columns=['_year'])
        dte = df[df['_year'] == Y].drop(columns=['_year'])
        r = _eval_fold_fixed(dtr, dte, horizon, kind, feats)
        if r is None:
            continue
        for m in models:
            per[m].append(r[f'{m}_auc']); per_ks[m].append(r[f'{m}_ks']); per_ic[m].append(r[f'{m}_ic'])
            per_ai[m].append(r[f'{m}_auc_incl']); per_ki[m].append(r[f'{m}_ks_incl'])
        print(f"  折{Y}: LGB {r['lgb_auc']:.3f}/{r['lgb_ks']:.3f}(含灰{r['lgb_auc_incl']:.3f}) | "
              f"LR {r['lr_auc']:.3f}/{r['lr_ks']:.3f}(含灰{r['lr_auc_incl']:.3f}) | "
              f"SC {r['sc_auc']:.3f}/{r['sc_ks']:.3f}(含灰{r['sc_auc_incl']:.3f})")
    print('\n' + '=' * 70)
    print(f'🔒 锁定特征 LOYO 结果(部署对齐):')
    out = {}
    for m, label in (('lgb', 'LGB'), ('lr', 'LR'), ('sc', 'SC 评分卡')):
        _a = np.array([x for x in per[m] if x == x], float)      # 丢 nan 折(样本不足/单类)
        _k = np.array([x for x in per_ks[m] if x == x], float)
        _i = np.array([x for x in per_ic[m] if x == x], float)
        _ai = np.array([x for x in per_ai[m] if x == x], float)
        _ki = np.array([x for x in per_ki[m] if x == x], float)
        print(f'  {label:<10} 不含灰 AUC {np.mean(_a):.3f}±{np.std(_a):.3f} | KS {np.mean(_k):.3f}±{np.std(_k):.3f} '
              f'| 含灰 AUC {np.mean(_ai):.3f}±{np.std(_ai):.3f} | KS {np.mean(_ki):.3f}±{np.std(_ki):.3f} '
              f'| IC {np.mean(_i):+.3f}±{np.std(_i):.3f}')
        out[m] = {'auc_mean': float(np.mean(_a)) if len(_a) else None,
                  'auc_std': float(np.std(_a)) if len(_a) else None,
                  'ks_mean': float(np.mean(_k)) if len(_k) else None,
                  'ks_std': float(np.std(_k)) if len(_k) else None,
                  'ic_mean': float(np.mean(_i)) if len(_i) else None,
                  'ic_std': float(np.std(_i)) if len(_i) else None,
                  'incl_auc_mean': float(np.mean(_ai)) if len(_ai) else None,
                  'incl_auc_std': float(np.std(_ai)) if len(_ai) else None,
                  'incl_ks_mean': float(np.mean(_ki)) if len(_ki) else None,
                  'incl_ks_std': float(np.std(_ki)) if len(_ki) else None,
                  'n_folds': int(len(_a)),
                  'auc_per_year': [None if a != a else round(float(a), 3) for a in per[m]]}
    return out


def main():
    ap = argparse.ArgumentParser(description='LOYO 评估(标准 / 锁定特征对齐验证)')
    ap.add_argument('features_path', help='features_derived.parquet')
    ap.add_argument('--fixed-features', default=None,
                    help='锁定特征(逗号分隔) → 跑部署对齐 LOYO(跳过每折选特征); 如 "个股PB,盈利能力_delta_1y"')
    ap.add_argument('--horizon', type=_parse_horizon, default=7)
    ap.add_argument('--kind', choices=['gray'], default='gray')
    args = ap.parse_args()
    if args.fixed_features:
        loyo_fixed(args.features_path, args.horizon, args.kind, args.fixed_features)
    else:
        run(args.features_path)


if __name__ == '__main__':
    main()
