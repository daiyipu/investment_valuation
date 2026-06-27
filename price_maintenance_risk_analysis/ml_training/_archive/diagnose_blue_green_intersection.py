#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""蓝绿线上模型交集诊断 —— AND 交集 / 概率平均 能否提升高概率区区分度。

复用 predict 的 SC 打分链路(apply_woe + lr_model.predict_proba), 对 features_derived
历史样本同时用 GREEN(current.full) 与 BLUE(同期限特征数差异最大的 gray SC)打分:
  1) 两模型预测独立性 Spearman ρ(决定交集收益的判决书);
  2) 全样本判别力: 非灰 AUC/KS + 全样本 IC(GREEN / BLUE / 平均);
  3) 高概率区(top X%)区分度: 单模型 / 交集AND / 平均 的 赢率·均收益·中位·样本量·95%CI。

注: 含训练样本(in-sample), 绝对值偏乐观; 相对对比(交集 vs 单模型)有效。
    严格 OOT 用 --oot-year 过滤(如 2024)。
"""
import os, sys, pickle, argparse
import numpy as np
import pandas as pd
from scipy.stats import spearmanr
from sklearn.metrics import roc_auc_score

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from deploy.db_model_store import load_predict_bundle, list_model_metas, get_model_meta
from deploy.model_registry import get_current
from eval_loyo import apply_woe
from validate_methods import calc_ks
from train.train_horizon_models import GRAY_CFG

PARQUET = os.path.join(HERE, 'data', 'features_derived.parquet')


def score_sc(sc, df):
    """SC 评分卡打分: median 填缺 → apply_woe → LR.predict_proba。"""
    feats = sc['features']
    med = sc.get('medians', {}) or {}
    X = pd.DataFrame(index=df.index)
    for f in feats:
        X[f] = pd.to_numeric(df[f], errors='coerce') if f in df.columns else np.nan
    X = X.fillna({f: med.get(f, 0) for f in feats}).replace([np.inf, -np.inf], 0)
    Xw = apply_woe(X, feats, sc['woe_bins']).replace([np.inf, -np.inf], 0).fillna(0)
    return sc['lr_model'].predict_proba(Xw)[:, 1]


def wilson(p, n, z=1.96):
    """比例的 Wilson 95% 置信区间。"""
    if n == 0:
        return (np.nan, np.nan)
    den = 1 + z**2 / n
    c = (p + z**2 / (2 * n)) / den
    h = z * np.sqrt(p * (1 - p) / n + z**2 / (4 * n**2)) / den
    return c - h, c + h


def pick_blue(green_ver, cur_nf, horizon):
    """复刻 predict 的 BLUE 选择: 同期限 gray SC, 排 current, 特征数差异最大。"""
    metas = list_model_metas(kind='gray')
    cands = [m for m in metas if m['version'] != green_ver and not m.get('lgb_model')
             and m.get('kind') == 'gray' and m.get('horizon') == horizon]
    cands.sort(key=lambda m: abs(m.get('n_features', 0) - cur_nf), reverse=True)
    return cands[0]['version'] if cands else None


def main():
    ap = argparse.ArgumentParser(description='蓝绿交集诊断(AND/平均 vs 单模型高区分度)')
    ap.add_argument('--frac', type=float, default=0.2, help='高区定义: top 分位(默认0.2=前20%)')
    ap.add_argument('--oot-year', type=int, default=0, help='只看该年及以后(0=全样本)')
    ap.add_argument('--blue', help='手动指定 BLUE 版本(默认按 predict 逻辑自动选)')
    args = ap.parse_args()

    df = pd.read_parquet(PARQUET).dropna(subset=['报价日']).reset_index(drop=True)
    if args.oot_year:
        yr = pd.to_numeric(df['报价日'], errors='coerce') // 10000
        df = df[yr >= args.oot_year].reset_index(drop=True)
    ret = pd.to_numeric(df['7个月涨跌幅'], errors='coerce')
    h = 7
    lo, hi = GRAY_CFG[h]
    print(f'7m gray 标签: 赢>{hi}% / 输<{lo}% | 样本 {len(df)} (ret有效 {int(ret.notna().sum())})')

    green_ver = get_current('full')
    sc_g = pickle.loads(load_predict_bundle(green_ver)['lr_bundle'])
    cur_h = (get_model_meta(green_ver) or {}).get('horizon', h)
    blue_ver = args.blue or pick_blue(green_ver, len(sc_g['features']), cur_h)
    print(f'GREEN: {green_ver} ({len(sc_g["features"])}feat)')
    print(f'BLUE : {blue_ver}')
    if not blue_ver:
        print('无 BLUE 候选, 退出'); return
    sc_b = pickle.loads(load_predict_bundle(blue_ver)['lr_bundle'])
    print(f'BLUE 特征数: {len(sc_b["features"])}')

    p_g = score_sc(sc_g, df)
    p_b = score_sc(sc_b, df)
    p_avg = (p_g + p_b) / 2
    rho = spearmanr(p_g, p_b).correlation
    print(f'\n=== 1. 独立性(决定交集收益) ===')
    print(f'  Spearman ρ(GREEN, BLUE) = {rho:.3f}'
          f'  → {">0.85 交集≈白做" if rho > 0.85 else "0.6-0.85 有收益但打折" if rho > 0.6 else "<0.6 收益最大"}')

    print(f'\n=== 2. 全样本判别力(非灰 AUC/KS + 全样本IC) ===')
    ng = (ret > hi) | (ret < lo)
    y = (ret[ng] > hi).astype(int)
    print(f'  {"模型":<10}{"AUC":>8}{"KS":>8}{"IC":>8}')
    for name, p in [('GREEN', p_g), ('BLUE', p_b), ('平均avg', p_avg)]:
        auc = roc_auc_score(y, p[ng])
        ks = calc_ks(y.values, p[ng])
        v = ret.notna().values
        ic = spearmanr(p[v], ret.values[v]).correlation
        print(f'  {name:<10}{auc:>8.3f}{ks:>8.3f}{ic:>8.3f}')

    def topmask(p, frac):
        return p >= np.quantile(p, 1 - frac)

    sets = [('GREEN单', topmask(p_g, args.frac)), ('BLUE单', topmask(p_b, args.frac)),
            ('交集AND', topmask(p_g, args.frac) & topmask(p_b, args.frac)),
            ('平均avg', topmask(p_avg, args.frac))]
    base_win = (ret[ret.notna()] > hi).mean()
    print(f'\n=== 3. 高概率区(top {args.frac*100:.0f}%)区分度 ===')
    print(f'  全体基线 赢率(>{hi}%): {base_win*100:.1f}% | 平均收益 {ret.mean():.2f}%\n')
    print(f'  {"区":<10}{"n":>6}{"赢率%":>8}{"赢率95%CI":>18}{"均收益%":>9}{"中位%":>8}')
    for name, m in sets:
        r = ret[m & ret.notna()]
        n = len(r)
        w = (r > hi).mean() if n else np.nan
        clo, chi = wilson(w, n)
        print(f'  {name:<10}{n:>6}{w*100:>8.1f}  [{clo*100:>4.1f},{chi*100:>4.1f}]'
              f'{r.mean():>9.2f}{r.median():>8.2f}')

    # 参考: 更极端 top 10%
    if args.frac > 0.1:
        print(f'\n=== (参考) top 10% ===')
        print(f'  {"区":<10}{"n":>6}{"赢率%":>8}{"赢率95%CI":>18}{"均收益%":>9}')
        for name, m in [('GREEN', topmask(p_g, 0.1)), ('BLUE', topmask(p_b, 0.1)),
                        ('交集AND', topmask(p_g, 0.1) & topmask(p_b, 0.1)),
                        ('平均avg', topmask(p_avg, 0.1))]:
            r = ret[m & ret.notna()]; n = len(r)
            w = (r > hi).mean() if n else np.nan
            clo, chi = wilson(w, n)
            print(f'  {name:<10}{n:>6}{w*100:>8.1f}  [{clo*100:>4.1f},{chi*100:>4.1f}]{r.mean():>9.2f}')


if __name__ == '__main__':
    main()
