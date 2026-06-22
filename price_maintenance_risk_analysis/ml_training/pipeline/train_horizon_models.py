#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""各期限独立模型: 每个期限/标签口径用【自己的】特征子集训练。

方法论(对应用户决策):
  1. 长短期 target 敏感指标不同 → 每期限独立选特征, 不共用 203 池。
  2. 特征选择 = IV top-30(train) → PSI 稳定性过滤(train vs test 分布漂移 >0.25 丢弃)。
     PSI 直接量 regime 漂移(我们一直头疼的), 比 top-K 重叠更对症。
  3. 覆盖 阈值-10 + 灰度(各期限调阈值) 共 10 个模型, 全部 set_current=False。

灰度阈值(语义约束 赢家>=+5%/输家<=-10%, 优选 n>=400 且正占比近 50%):
  1m[-10,+10] 3m[-15,+10] 6m[-20,+10] 7m[-20,+5] 12m[-20,+5]

用法: python train_horizon_models.py <features_derived.parquet> [--split-year 2024]
特征选择阈值固定在 feature_selection.py(N_IV/PSI_MAX/CORR_MAX/VIF_MAX), 不随运行变。
输出: output/horizon_models_comparison.csv + output/per_horizon_features/*.csv + registry(10 条)
"""
import argparse, os, sys, json, pickle
import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from validate_methods import make_features, eval_metrics
from feature_exclusions import get_excluded_columns
from train_models import LGB_PARAMS
# 特征选择统一走标准模块 feature_selection(IV→PSI→去相关→VIF→LGBM), 不再各脚本各搞一套
from feature_selection import (select_features, prune_by_lgb_importance,
                               pipeline_summary, IV_MIN, PSI_MAX, CORR_MAX, VIF_MAX)
from db_model_store import save_model_meta   # 模型元信息入 DB(替代散落 meta.json)

HORIZONS = [1, 3, 6, 7, 12]   # 月期限(批量 eval_loyo.run 用); 周 '1w'/'2w' 走 train_to_production 单期限路径
# 灰度区阈值(lose,win)。月: 1m/3m sweep 定 (-10,10); 6/7/12m 沿用 [-20,+10]。
# 周: 2026-06-22 sweep_label 定(全样本IC最优); 周收益波动小于月→阈值更紧(±5/±6)。
GRAY_CFG = {1: (-10, 10), 3: (-10, 10), 6: (-20, 10), 7: (-20, 10), 12: (-20, 10),
            '1w': (-5, 5), '2w': (-6, 6)}


def _ret_col(horizon):
    """返回列名: 月用 int(→'{h}个月涨跌幅'), 周用 str 'Nw'(→'N周涨跌幅')。"""
    if isinstance(horizon, str) and horizon.endswith('w'):
        return f'{horizon[:-1]}周涨跌幅'
    return f'{horizon}个月涨跌幅'


def _tag(horizon):
    """版本/标签 tag: 7→'7m', '1w'→'1w'。"""
    return horizon if isinstance(horizon, str) else f'{horizon}m'


def _parse_horizon(s):
    """CLI --horizon 解析: '7'→7(int月), '1w'→'1w'(str周)。"""
    s = str(s).strip()
    return int(s) if s.lstrip('-').isdigit() else s


def _prep(X_raw, medians=None):
    excl = [c for c in get_excluded_columns(X_raw.columns) if c in X_raw.columns]
    X = X_raw.drop(columns=excl)
    if medians is None:
        medians = X.median()
    return X.fillna(medians).replace([np.inf, -np.inf], 0), medians


def _train(X, y):
    import lightgbm as lgb
    from sklearn.linear_model import LogisticRegression
    from sklearn.model_selection import train_test_split
    from sklearn.preprocessing import StandardScaler
    Xt, Xv, yt, yv = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)
    gbm = lgb.LGBMClassifier(**LGB_PARAMS)
    gbm.fit(Xt, yt, eval_set=[(Xv, yv)], callbacks=[lgb.early_stopping(20, verbose=False)])
    sc = StandardScaler().fit(X)
    lr = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)
    lr.fit(sc.transform(X), y)
    return gbm, lr, sc


def build_label(df, horizon, kind='gray'):
    """返回 (label_col, gray_cfg)。gray 实时构造, 列名带「标签」前缀防泄漏。
    horizon: int=月(1/3/6/7/12) 或 str='1w'/'2w'(周)。thr 旧二分类已移除(2026-06-22), 仅 gray。"""
    lo, hi = GRAY_CFG[horizon]
    col = f'标签_极性_灰度自定义_{lo}_{hi}_{_tag(horizon)}'
    ret = pd.to_numeric(df[_ret_col(horizon)], errors='coerce')
    df[col] = np.where(ret > hi, 1, np.where(ret < lo, 0, np.nan))
    return col, (lo, hi)


def run(features_path, split_year=2024):
    from model_registry import register_version
    raw = pd.read_parquet(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    raw['_y'] = (pd.to_numeric(raw['报价日'], errors='coerce') // 10000).astype('Int64')
    dtr = raw[raw['_y'] <= split_year].drop(columns=['_y'])
    dte = raw[raw['_y'] >= split_year + 1].drop(columns=['_y'])

    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'output')
    feat_dir = os.path.join(out_dir, 'per_horizon_features')
    os.makedirs(feat_dir, exist_ok=True)

    rows, today = [], pd.Timestamp.now().strftime('%Y%m%d_%H%M')
    derived_ver = 'derived_20260616_1627_a77c2bb4_7m'   # 当前 features_derived 快照

    for h in HORIZONS:
        for kind in ('gray',):
            lbl, gcfg = build_label(dtr, h, kind)   # 注意: gray 在 dtr 上注入列; dte 需同样注入
            if kind == 'gray':
                build_label(dte, h, kind)
            Xtr_raw, ytr, _ = make_features(dtr, label_col=lbl, ret_col=f'{h}个月涨跌幅')
            Xte_raw, yte, _ = make_features(dte, label_col=lbl, ret_col=f'{h}个月涨跌幅')
            if Xtr_raw is None or len(ytr) < 40 or ytr.nunique() < 2 or Xte_raw is None or yte.nunique() < 2:
                print(f'  ⚠ {h}m/{kind}: 样本不足, 跳过'); continue
            Xtr, med = _prep(Xtr_raw)
            Xte = _prep(Xte_raw, medians=med)[0].reindex(columns=Xtr.columns)
            # 标准五步之 1-4: IV→PSI→去相关→VIF (阈值固定在 feature_selection)
            kept_vif, fdf = select_features(Xtr, ytr, Xte, Xtr_raw=Xtr_raw, Xte_raw=Xte_raw)
            gbm, lr, sc = _train(Xtr[kept_vif], ytr)
            fdf['lgb_imp'] = fdf['feature'].map(
                dict(zip(kept_vif, gbm.feature_importances_))).fillna(0).astype(int)
            # 步5: LGBM 重要性剪枝(剔树未分裂的特征), 剪后重训
            kept = prune_by_lgb_importance(gbm, kept_vif)
            if 5 <= len(kept) < len(kept_vif):
                gbm, lr, sc = _train(Xtr[kept], ytr)
            else:
                kept = list(kept_vif)
            Xtr_s, Xte_s = Xtr[kept], Xte[kept]
            p_lgb = gbm.predict_proba(Xte_s)[:, 1]
            p_lr = lr.predict_proba(sc.transform(Xte_s))[:, 1]
            al, ar = eval_metrics(yte.values, p_lgb), eval_metrics(yte.values, p_lr)
            sel_summary = pipeline_summary(kept_vif, kept, fdf)

            # 灰度: 全量 3 桶(输家<lo / 灰[lo,hi] / 赢家>hi)打分, 看分数跨度
            band_spread = np.nan
            if kind == 'gray':
                ret_full = pd.to_numeric(raw[f'{h}个月涨跌幅'], errors='coerce').reset_index(drop=True)
                raw_clean = raw.drop(columns=['_y'], errors='ignore').reset_index(drop=True)
                Xf, _ = _prep(raw_clean, medians=med)
                Xf_s = Xf.reindex(columns=kept).fillna(med)
                pf = gbm.predict_proba(Xf_s)[:, 1]
                lo, hi = gcfg
                bucket = np.where(ret_full < lo, '输', np.where(ret_full > hi, '赢', '灰'))
                grp = pd.Series(pf).groupby(bucket).mean()
                band_spread = round(float(grp.max() - grp.min()), 3)

            cfg_tag = f'{h}m_{"thr" if kind=="thr" else "gray"}'
            ver = f'v_{today}_{cfg_tag}_{len(kept)}feat_auc{al["auc"]:.2f}'.replace('.', '')
            # 特征清单(每特征 IV/PSI/相关/VIF/LGBM重要, 供审计) 存 CSV; 这不是 meta
            fdf.to_csv(os.path.join(feat_dir, f'{cfg_tag}_features.csv'), index=False)
            # 模型 + 元信息 + 权重 全部入 DB(替代散落的 version 目录: meta.json/lgb.txt/lr.pkl)
            meta = {
                'version': ver, 'label_config': cfg_tag, 'kind': kind, 'horizon': h,
                'gray_cfg': gcfg, 'features': kept, 'n_features': len(kept),
                'selection': sel_summary,
                'selection_thresholds': {'iv_min': IV_MIN, 'psi_max': PSI_MAX,
                                         'corr_max': CORR_MAX, 'vif_max': VIF_MAX},
                'metrics': {'lgb_oot_auc': al['auc'], 'lgb_oot_ks': al['ks'],
                            'lr_oot_auc': ar['auc'], 'lr_oot_ks': ar['ks'],
                            'band_spread': band_spread, 'split_year': split_year},
                'medians': {f: float(med[f]) for f in kept},          # predict 填缺失用
                'lgb_model': gbm.booster_.model_to_string(),          # LGB 权重(文本)
                'lr_bundle': pickle.dumps({'model': lr, 'scaler': sc, 'features': kept}),  # LR 权重
                'n_train': int(len(ytr)), 'n_test': int(len(yte)),
                'train_pos': float(ytr.mean()), 'test_pos': float(yte.mean()),
                'dataset_version': derived_ver,
                'note': f'标准五步特征选择({sel_summary})',
            }
            try:
                save_model_meta(meta)
            except Exception as e:
                print(f'    ⚠ 模型入 DB 失败(权重未存!): {e}')
            register_version('full', ver, ver,                      # dir=ver 占位(权重在 DB, 无磁盘目录)
                             metrics={'lgb_oot_auc': al['auc'], 'lgb_oot_ks': al['ks'],
                                      'lr_oot_auc': ar['auc'], 'lr_oot_ks': ar['ks'],
                                      'band_spread': band_spread},
                             n_features=len(kept), threshold=-10,
                             n_samples=len(ytr), positive_rate=float(ytr.mean()),
                             files=['(in ml_model_meta DB)'],
                             note=f'标准五步特征选择({sel_summary})',
                             set_current=False,
                             label_config=cfg_tag, dataset_version=derived_ver)

            rows.append({'cfg': cfg_tag, 'horizon': f'{h}m', 'kind': kind,
                         'gray_cfg': (f'{gcfg[0]}/{gcfg[1]}' if gcfg else '-'),
                         'n_train': len(ytr), 'train_pos': round(ytr.mean()*100),
                         'n_test': len(yte), 'test_pos': round(yte.mean()*100),
                         'selection': sel_summary, 'n_feat': len(kept),
                         'lgb_auc': round(al['auc'], 3), 'lgb_ks': round(al['ks'], 3),
                         'lr_auc': round(ar['auc'], 3), 'lr_ks': round(ar['ks'], 3),
                         'band_spread': band_spread, 'version': ver})
            print(f"  {cfg_tag:<10} [{sel_summary}] train {len(ytr)}({ytr.mean()*100:.0f}%) "
                  f"test {len(yte)} | LGB {al['auc']:.3f}/{al['ks']:.3f} "
                  f"LR {ar['auc']:.3f}/{ar['ks']:.3f}"
                  + (f" | 桶跨度{band_spread}" if kind == 'gray' else ''))

    out = pd.DataFrame(rows)
    path = os.path.join(out_dir, 'horizon_models_comparison.csv')
    out.to_csv(path, index=False)
    pd.set_option('display.width', 200); pd.set_option('display.max_columns', 20)
    print('\n' + '=' * 90)
    print(out.to_string(index=False))
    print(f'\n写出: {path} | 特征清单: {feat_dir}/ | 全部 set_current=False')


def main():
    ap = argparse.ArgumentParser(description='各期限独立特征集模型训练')
    ap.add_argument('features_path', help='features_derived.parquet')
    ap.add_argument('--split-year', type=int, default=2024)
    args = ap.parse_args()
    run(args.features_path, split_year=args.split_year)


if __name__ == '__main__':
    main()
