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
import warnings

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # train/→ml_training/→PKG
for _p in (PKG, os.path.join(PKG,'ml_training'), os.path.join(PKG,'ml_training','pipeline'), os.path.join(PKG,'scripts')):
    if _p not in sys.path: sys.path.insert(0, _p)
from validate.save_validation_db import load_features  # noqa: E402  (DB样本空间直读)
from validate.validate_methods import make_features, eval_metrics
from features.feature_selection import select_features, pipeline_summary, IV_MIN, PSI_MAX, CORR_MAX, VIF_MAX
from train.train_horizon_models import GRAY_CFG, build_label, _prep, _ret_col, _tag, _parse_horizon
from validate.eval_loyo import fit_woe, apply_woe
from sklearn.tree import DecisionTreeClassifier, export_text
from deploy.db_model_store import save_model_meta
from deploy.model_registry import register_version
from sklearn.linear_model import LogisticRegression
from utils.db_manager import ValuationDB

WOE_FILL = lambda X: X.replace([np.inf, -np.inf], np.nan).fillna(0)


def scorecardpy_binning(X, y, features, method='chimerge', max_bins=5):
    """使用scorecardpy进行真正的卡方分箱或决策树分箱"""
    import scorecardpy as sc
    import warnings
    warnings.filterwarnings('ignore')

    print(f"  === scorecardpy分箱 ===")

    # 准备scorecardpy格式的数据（标签列名必须为'y'）
    df = X[features].copy()
    df['y'] = y.values

    # 🔧 关键修复: 转换y列为int类型以兼容pandas 2.3.3
    df['y'] = df['y'].astype(int)

    print(f"  数据形状: {df.shape}, 特征数: {len(features)}")
    print(f"  y列类型: {df['y'].dtype} (修复后)")

    # 检查数据质量
    if df.isnull().sum().sum() > 0:
        print(f"  ⚠️ 数据存在缺失值: {df.isnull().sum().sum()}，尝试填充...")
        for col in features:
            if df[col].isnull().sum() > 0:
                df[col] = df[col].fillna(df[col].median())

    wbins = {}
    try:
        print(f"  调用scorecardpy {method}分箱...")
        bins = sc.woebin(df, y='y', method=method, max_bin=max_bins)
        print(f"  ✅ scorecardpy分箱成功!")

        # 转换scorecardpy分箱结果到我们的格式
        for f in features:
            if f in bins:
                bin_data = bins[f]
                rights = []
                woes = []

                for _, row in bin_data.iterrows():
                    bin_str = row['bin']
                    # 处理分箱边界和WOE值
                    try:
                        # 提取WOE值
                        woe_val = float(row['woe'])

                        # 解析边界
                        if 'inf' in bin_str or 'inf' in str(row.values):
                            # 处理无限边界情况
                            woes.append(woe_val)
                            continue

                        if ',' in bin_str:
                            parts = bin_str.strip('[]()').split(',')
                            right_val = float(parts[1].strip(')'))
                            rights.append(right_val)
                            woes.append(woe_val)
                        else:
                            # 单一边界情况
                            woes.append(woe_val)
                    except:
                        continue

                # 确保至少有一个分箱边界
                if len(woes) > 0:
                    # 如果没有rights（即单箱情况），设置一个虚拟边界
                    if len(rights) == 0:
                        # 单箱情况：设置一个极大边界，所有值都在第一箱
                        rights = [float('inf')]
                        # 单箱情况的woe值重复
                        if len(woes) == 1:
                            woes = [woes[0], woes[0]]

                    wbins[f] = {'rights': rights, 'woes': woes}
                    print(f"    {f}: {len(rights)+1} 个分箱 (边界: {len(rights)})")
                else:
                    print(f"    {f}: 无有效分箱，跳过")

    except Exception as e:
        print(f"  ❌ scorecardpy分箱失败: {e}")
        print(f"  错误类型: {type(e).__name__}")

        # 详细诊断
        import traceback
        print("  详细错误信息:")
        traceback.print_exc()
        return {}

    return wbins


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


def run(features_path, horizon, kind, split_year, set_current, features=None, loyo_stats=None, use_mysql=False, sample_size=10000, binning_method="tree"):
    if use_mysql:
        import pymysql
        print("🔗 从MySQL宽表加载定增样本...")
        conn = pymysql.connect(**ValuationDB.MYSQL_CONFIG)
        query = f"SELECT * FROM ml_features_wide WHERE `定增决策` IS NOT NULL AND `{_ret_col(horizon)}` IS NOT NULL ORDER BY RAND() LIMIT {sample_size}"
        df = pd.read_sql(query, conn)
        conn.close()
        print(f"✅ MySQL定增数据加载成功: {len(df)}样本")
    else:
        df = load_features(features_path).dropna(subset=['报价日']).reset_index(drop=True)
    df['_y'] = pd.to_datetime(df['报价日'], errors='coerce').dt.year.astype('Int64')
    print(f"原始数据年份分布: {df['_y'].value_counts().sort_index().to_dict()}")

    dtr_s = df[df['_y'] <= split_year].drop(columns=['_y'])      # 选特征用(算 PSI)
    dte_s = df[df['_y'] >= split_year + 1].drop(columns=['_y'])
    print(f"训练集: {len(dtr_s)}, 验证集: {len(dte_s)}")
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

    # 使用传统分箱方法，但加入单调性检查和波浪形特征优化
    print(f"使用{binning_method}分箱方法...")

    # 修复索引问题，确保Xall和yall索引一致
    Xall_fixed = Xall[kept].reset_index(drop=True)
    yall_fixed = yall.reset_index(drop=True)

    if binning_method in ['chimerge', 'tree']:
        # 使用真正的scorecardpy分箱
        wbins = scorecardpy_binning(Xall_fixed, yall_fixed, kept, method=binning_method, max_bins=5)
        if wbins:
            # scorecardpy分箱成功，应用WOE变换
            Xall_w = apply_woe(Xall_fixed, kept, wbins)
        else:
            # scorecardpy分箱失败，回退到传统方法
            print("scorecardpy分箱失败，回退到传统方法...")
            Xall_w, wbins = fit_woe(Xall_fixed, yall_fixed, kept)
    else:
        # 使用传统分箱方法
        Xall_w, wbins = fit_woe(Xall_fixed, yall_fixed, kept)

    # 如果使用scorecardpy分箱且成功，跳过波浪形特征优化
    if binning_method not in ['chimerge', 'tree'] or not wbins:
        # 检测波浪形特征
        print("检测波浪形特征...")
        wave_features = []
        for f in kept:
            if f in wbins:
                woes = wbins[f]['woes']
                # 检查单调性
                changes = sum(1 for i in range(1, len(woes)) if woes[i] < woes[i-1])
                if changes > 1:  # 超过一次趋势变化
                    wave_features.append(f)
                    print(f"  {f}: 波浪形({changes}次趋势变化)")

        if wave_features:
            print(f"发现{len(wave_features)}个波浪形特征，进行合并优化...")
            # 对波浪形特征进行合并优化
            for f in wave_features:
                if f in wbins:
                    rights = wbins[f]['rights']
                    woes = wbins[f]['woes']
                    # 简单合并策略：合并相邻的相同趋势分箱
                    if len(rights) > 3:
                        # 合并中间分箱
                        new_rights = [rights[0], rights[-2]]
                        # 重新计算WOE
                        feature_values = Xall_fixed[f].values
                        y_values = yall_fixed.values
                        new_woes = []
                        for i in range(len(new_rights) + 1):
                            if i == 0:
                                mask = feature_values <= new_rights[0]
                            elif i == len(new_rights):
                                mask = feature_values > new_rights[-1]
                            else:
                                mask = (feature_values > new_rights[i-1]) & (feature_values <= new_rights[i])

                            if mask.sum() > 0:
                                good_rate = y_values[mask].mean()
                                bad_rate = 1 - good_rate
                                if bad_rate > 0 and good_rate > 0:
                                    woe = np.log(good_rate / bad_rate)
                                else:
                                    woe = 0
                                new_woes.append(woe)
                            else:
                                new_woes.append(0)

                        wbins[f] = {'rights': new_rights, 'woes': new_woes}
                        print(f"  {f}: 从{len(rights)}箱合并到{len(new_rights)}箱")

    lr = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)
    lr.fit(WOE_FILL(Xall_w), yall_fixed)

    # 全量训练后自评: test 年已入训练 → 含泄漏偏高, 非真实泛化(仅拟合参考)。
    # 真泛化能力看 loyo_stats(由 train_to_production 跑 loyo_fixed 传入)。
    print(f"验证集特征: {Xte_s[kept].shape}")
    Xte_s_w = apply_woe(Xte_s[kept], kept, wbins)
    print(f"WOE变换后验证集: {Xte_s_w.shape}")

    if Xte_s_w.shape[0] == 0:
        print("❌ 验证集WOE变换后样本数为0，跳过自评")
        ate = {'auc': None, 'ks': None}
    else:
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
    ap.add_argument('features_path', nargs='?', help='特征文件路径(MySQL模式下可选)')
    ap.add_argument('--horizon', type=_parse_horizon, default=7)
    ap.add_argument('--kind', choices=['gray'], default='gray')
    ap.add_argument('--split-year', type=int, default=2024)
    ap.add_argument('--set-current', action='store_true', help='设为 current.full(生产)')
    ap.add_argument('--features', default=None, help='锁定特征(逗号分隔, 跳过select_features); 如共识特征')
    ap.add_argument('--use-mysql', action='store_true', help='使用MySQL宽表数据')
    ap.add_argument('--sample-size', type=int, default=10000, help='MySQL采样数量')
    ap.add_argument('--binning-method', choices=['tree', 'chimerge'], default='tree',
                    help='分箱方法: tree=决策树分箱, chimerge=卡方分箱')
    args = ap.parse_args()
    feats = args.features.split(',') if args.features else None

    # MySQL模式下features_path可选
    if args.use_mysql:
        features_path = args.features_path or "mysql_data"
    else:
        if not args.features_path:
            ap.error('features_path is required when not using --use-mysql')
        features_path = args.features_path

    run(features_path, args.horizon, args.kind, args.split_year, args.set_current, features=feats,
        use_mysql=args.use_mysql, sample_size=args.sample_size, binning_method=args.binning_method)


if __name__ == '__main__':
    main()
