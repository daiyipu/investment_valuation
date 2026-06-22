#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""标准"从特征到生产"流水线(共识特征规则, 2026-06-17 固化)。

为什么: 标准五步选特征跨折不稳(6 折常只共有个别特征), 部署用某折的"一次抽取"
会和验证偏离。共识特征(跨折反复出现的)更稳、方差更小、可解释 → 作生产特征集。

流程(加完新特征后跑这一条即可):
  1. 各 LOYO 折 select_features → 特征跨折频次
  2. 共识 = 频次 ≥ --min-folds 的特征(跨期稳定的根本面; 太少则自动放宽)
  3. 锁定共识特征 LOYO(LGB/LR/SC) → 与部署对齐的真实性能(mean±std)
  4. 全量训最终模型(--model sc|lgb), 入库 set_current, 打印评分卡(SC)

用法:
  python train_to_production.py <features_derived.parquet> --horizon 7 --kind gray \
      [--min-folds 3] [--model sc] [--set-current]
"""
import argparse
import os
import sys
import pickle
import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_methods import make_features
from feature_selection import select_features, CORR_MAX, VIF_MAX
from train_horizon_models import GRAY_CFG, build_label, _prep, _train, _ret_col, _tag, _parse_horizon
from train_scorecard import calc_iv_all_features, remove_correlated, filter_by_vif
from eval_loyo import loyo_fixed, fit_woe, apply_woe
from db_model_store import save_model_meta
from model_registry import register_version
from train_scorecard_model import print_scorecard, WOE_FILL, run as run_sc


def derive_consensus(df, horizon, kind, min_folds):
    """各 LOYO 折 select_features(真实 held-out 年判 PSI) → 特征频次 → 共识(频次≥min_folds)。"""
    df = df.dropna(subset=['报价日']).reset_index(drop=True)
    df['_y'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000).astype('Int64')
    years = sorted(int(y) for y in df['_y'].dropna().unique())
    from collections import Counter
    freq = Counter()
    ret = _ret_col(horizon)
    print(f'推导共识特征: {len(years)} 折各跑 select_features(真实 held-out 年判 PSI)...')
    for Y in years:
        dtr = df[df['_y'] != Y].drop(columns=['_y'])
        dte = df[df['_y'] == Y].drop(columns=['_y'])
        lbl, _ = build_label(dtr, horizon, kind)
        if kind == 'gray':
            build_label(dte, horizon, kind)
        Xtr_raw, ytr, _ = make_features(dtr, label_col=lbl, ret_col=ret)
        Xte_raw, yte, _ = make_features(dte, label_col=lbl, ret_col=ret)
        if Xtr_raw is None or Xte_raw is None:
            continue
        Xtr, med = _prep(Xtr_raw)
        Xte = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr.columns)
        kept, _ = select_features(Xtr, ytr, Xte, Xtr_raw=Xtr_raw, Xte_raw=Xte_raw)
        for f in kept:
            freq[f] += 1
        print(f'  折{Y} 选 {len(kept)}: {kept}')
    # 频次表
    print(f'\n特征跨折频次(共 {len(years)} 折):')
    for f, c in freq.most_common():
        print(f'  {f:<22} {c}/{len(years)} {"★" if c >= min_folds else ""}')
    consensus = [f for f, c in freq.most_common() if c >= min_folds]
    # 太少则放宽到 min_folds-1 (至少 4 个)
    mf = min_folds
    while len(consensus) < 4 and mf > 1:
        mf -= 1
        consensus = [f for f, c in freq.most_common() if c >= mf]
        print(f'  ⚠ 共识<4, 放宽 min_folds→{mf}')
    print(f'\n共识特征(min_folds={mf}, {len(consensus)}个): {consensus}')

    # ── 共识层去相关+VIF(补漏: 单折 select_features 各自去重, 但跨折聚合时 IV 排名翻转
    # 会让相关对的两个成员各在不同折入选 → 双双进共识, 如 ROE/ROE摊薄、净利增长/扣非净利增长)。
    # 用 pooled 全量数据再跑一次 remove_correlated+filter_by_vif, 塌掉跨折聚合产生的冗余。
    import contextlib, io
    df_pool = df.drop(columns=['_y'])
    lbl_pool, _ = build_label(df_pool, horizon, kind)
    Xpool_raw, ypool, _ = make_features(df_pool, label_col=lbl_pool, ret_col=_ret_col(horizon))
    if Xpool_raw is not None and len(consensus) > 1:
        Xpool, _ = _prep(Xpool_raw)
        feats = [f for f in consensus if f in Xpool.columns]
        with contextlib.redirect_stdout(io.StringIO()):
            iv_pool = calc_iv_all_features(Xpool_raw, ypool)
            dec = remove_correlated(Xpool, feats, iv_pool, threshold=CORR_MAX)
            dec = filter_by_vif(Xpool, dec, max_vif=VIF_MAX)
        dropped = [f for f in consensus if f not in dec]
        if len(dec) >= 4:
            consensus = dec
            if dropped:
                print(f'⚠ 共识层去相关/VIF 剔除 {len(dropped)} 个跨折冗余: {dropped}')
        else:
            print(f'⚠ 共识层去重后仅 {len(dec)} 个(<4), 保留去重前共识')
    # LOO 精修候选: 出现过但未达共识的特征(供前向 LOO 测试是否值得加回)
    candidates = [f for f, c in freq.most_common() if f not in consensus][:20]
    return consensus, candidates


def loo_refine(features_path, horizon, kind, consensus, candidates, lam=0.5, th=0.001):
    """LOO 精修(标准 scorecard ⑦VIF 之后的封装式精修补充): 后向删有害 + 前向加有益,
    按 LOYO score = mean_AUC − λ·std_AUC。迭代式(每步相对当前基线重测)。"""
    import contextlib, io

    def _sc(feats):
        feats = [f for f in feats if f]
        if len(feats) < 3:
            return None
        with contextlib.redirect_stdout(io.StringIO()):
            out = loyo_fixed(features_path, horizon, kind, ','.join(feats))
        return out['sc'] if out and out.get('sc') else None

    def _score(s):
        return (s['auc_mean'] - lam * s['auc_std']) if s else -9

    refined = [f for f in consensus]
    base = _sc(refined)
    if not base:
        return refined
    print(f'\n[LOO精修] 基线({len(refined)}feat): AUC {base["auc_mean"]:.4f}±{base["auc_std"]:.4f} '
          f'score={_score(base):.4f}')
    # 后向: 删有害+删中性(删了 score 不明显降 = 冗余/死重), 只留明确有贡献的。
    # 按 pooled IV 升序删(冗余对里低IV先删、高IV留存代表); 带 10 个下限防过度精简。
    df0 = pd.read_parquet(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    lbl0, _ = build_label(df0, horizon, kind)
    Xraw0, y0, _ = make_features(df0, label_col=lbl0, ret_col=_ret_col(horizon))
    iv = calc_iv_all_features(Xraw0, y0).set_index('feature')['iv']
    order = sorted([f for f in consensus if f in iv.index], key=lambda f: iv.get(f, 0))
    for f in order:
        if len(refined) <= 10:
            print(f'  达精简下限10, 停止后向')
            break
        sub = [x for x in refined if x != f]
        r = _sc(sub)
        if r and _score(r) >= _score(base) - th:  # 删f 后 score 不明显降 → 冗余/中性 → 删
            print(f'  LOO删 {f}(IV={iv.get(f,0):.3f}, 中性/冗余): score {_score(base):.4f}→{_score(r):.4f}')
            refined = sub
            base = r
    # 前向: 加有益(加了 score 上升)
    for f in candidates:
        if f in refined:
            continue
        r = _sc(refined + [f])
        if r and _score(r) > _score(base) + th:
            print(f'  LOO加 {f}: score {_score(base):.4f}→{_score(r):.4f}')
            refined = refined + [f]
            base = r
    print(f'[LOO精修] 结果({len(refined)}feat): AUC {base["auc_mean"]:.4f}±{base["auc_std"]:.4f}')
    return refined


def final_collapse(features_path, horizon, kind, refined, corr_th=0.9):
    """最终 corr 复核: 塌掉 LOO 重新引入的近重复(|r|>corr_th)。
    用严格阈值 0.9(只塌近重复), 避免 VIF(>0.7)那种激进误删 LOO 救回的有用特征。"""
    import contextlib, io
    df = pd.read_parquet(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    lbl, _ = build_label(df, horizon, kind)
    Xraw, y, _ = make_features(df, label_col=lbl, ret_col=_ret_col(horizon))
    X, _ = _prep(Xraw)
    feats = [f for f in refined if f in X.columns]
    iv = calc_iv_all_features(Xraw, y)
    with contextlib.redirect_stdout(io.StringIO()):
        dec = remove_correlated(X, feats, iv, threshold=corr_th)
    dropped = [f for f in feats if f not in dec]
    if dropped:
        print(f'[最终corr复核] 塌近重复(|r|>{corr_th}) {len(dropped)}: {dropped} → 保留 {len(dec)}')
    else:
        print(f'[最终corr复核] 无近重复, 保留 {len(feats)}')
    return dec


def deploy_lgb(features_path, horizon, kind, consensus, split_year, set_current):
    """LGB 部署: 全量训 LGB+LR(共识特征), 入库。"""
    df = pd.read_parquet(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    lbl, gcfg = build_label(df, horizon, kind)
    Xall_raw, yall, _ = make_features(df, label_col=lbl, ret_col=_ret_col(horizon))
    Xall, med = _prep(Xall_raw)
    feats = [f for f in consensus if f in Xall.columns]
    gbm, lr, sc = _train(Xall[feats], yall)
    # LGB 训练集概率 10 分位边界(部署后 predict 映射固定档位)
    train_proba = gbm.predict_proba(Xall[feats])[:, 1]
    proba_deciles = np.quantile(train_proba, np.linspace(0.1, 0.9, 9)).tolist()
    ver = f'v_lgb_{pd.Timestamp.now().strftime("%Y%m%d_%H%M")}_{_tag(horizon)}_{kind}_{len(feats)}feat'
    save_model_meta({'version': ver, 'label_config': f'{horizon}m_{kind}_lgb_consensus',
                     'kind': kind, 'horizon': horizon, 'gray_cfg': gcfg, 'features': feats,
                     'n_features': len(feats), 'medians': {f: float(med[f]) for f in feats},
                     'lgb_model': gbm.booster_.model_to_string(),
                     'lr_bundle': pickle.dumps({'model': lr, 'scaler': sc, 'features': feats,
                                                 'proba_deciles': proba_deciles}),
                     'metrics': {}, 'note': f'LGB 共识特征 {feats}',
                     'dataset_version': 'derived_20260616_2334_f35ba6f3_7m'})
    register_version('full', ver, ver, metrics={}, n_features=len(feats), threshold=-10,
                     n_samples=len(yall), positive_rate=float(yall.mean()),
                     files=['(in DB)'], note=f'LGB共识{feats}', set_current=set_current,
                     label_config=f'{_tag(horizon)}_{kind}_lgb_consensus')
    print(f'\n✅ LGB 入库 {ver} | {"已设生产" if set_current else "未切生产"}')
    return ver


def main():
    ap = argparse.ArgumentParser(description='共识特征→生产 标准流水线')
    ap.add_argument('features_path')
    ap.add_argument('--horizon', type=_parse_horizon, default=7)
    ap.add_argument('--kind', choices=['gray'], default='gray')
    ap.add_argument('--min-folds', type=int, default=3, help='共识阈值: 跨折出现≥此数(默认3/6)')
    ap.add_argument('--model', choices=['sc', 'lgb'], default='sc')
    ap.add_argument('--set-current', action='store_true')
    args = ap.parse_args()

    df = pd.read_parquet(args.features_path)
    consensus, candidates = derive_consensus(df, args.horizon, args.kind, args.min_folds)
    # LOO 精修 + 最终 corr 复核(标准 ⑦VIF之后、⑧LR之前的封装式精修)
    refined = loo_refine(args.features_path, args.horizon, args.kind, consensus, candidates)
    final = final_collapse(args.features_path, args.horizon, args.kind, refined)
    # 锁定精修特征 LOYO(对齐验证); SC 去偏 AUC/KS → 写模型 meta
    loyo_res = loyo_fixed(args.features_path, args.horizon, args.kind, ','.join(final))
    _sc_loyo = (loyo_res or {}).get('sc')
    # 部署
    if args.model == 'sc':
        print(f'\n--- 训 SC(精修特征) 并部署 ---')
        ver = run_sc(args.features_path, args.horizon, args.kind, 2024, args.set_current,
                     features=final, loyo_stats=_sc_loyo)
    else:
        ver = deploy_lgb(args.features_path, args.horizon, args.kind, final, 2024, args.set_current)

    # 报告(SC): 管线终步自动出 per-year AUC/KS/IC md(固化: 说标签→一键出模型+报告)
    if args.model == 'sc':
        from report_horizon import generate_report
        generate_report(args.features_path, args.horizon, ver)


if __name__ == '__main__':
    main()
