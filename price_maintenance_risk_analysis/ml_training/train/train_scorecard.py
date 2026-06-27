#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增盈利预测 - 评分卡模型训练

特征筛选流程:
  Step 1: IV (Information Value) 计算 → 初筛
  Step 2: LightGBM 重要性交叉验证
  Step 3: 去相关 (|r| > 0.7 的保留 IV 更高的)
  Step 4: 取 Top 10-15 特征

评分卡训练:
  WOE 分箱变换 → 逻辑回归 → 得分表 (Base=600, PDO=20)

用法:
    python ml_training/train_scorecard.py data/features_derived.parquet [--threshold -10] [--n-features 12]
"""

import sys
import os
import argparse
import pickle
import json
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')


# ====== Step 1: IV 计算 ======

def calc_iv_for_feature(series, y, n_bins=5):
    """计算单个特征的 WOE 和 IV

    Args:
        series: 数值型 Series
        y: 二分类标签 Series (0/1)
        n_bins: 分箱数

    Returns:
        dict: {iv, bins_info, coverage} 或 {iv: 0, skip: True}
    """
    valid = series.notna() & y.notna()
    s = series[valid].copy()
    y_valid = y[valid].copy()

    if len(s) < 50:
        return {'iv': 0, 'skip': True, 'reason': '样本不足50'}

    # 量化分箱
    try:
        bins = pd.qcut(s, n_bins, duplicates='drop')
    except ValueError:
        return {'iv': 0, 'skip': True, 'reason': '分箱失败'}

    df = pd.DataFrame({'bin': bins, 'y': y_valid.values})
    grouped = df.groupby('bin', observed=True)

    total_pos = df['y'].sum()
    total_neg = len(df) - total_pos

    if total_pos == 0 or total_neg == 0:
        return {'iv': 0, 'skip': True, 'reason': '单类别'}

    bins_info = []
    iv_total = 0

    for name, group in grouped:
        pos = group['y'].sum()
        neg = len(group) - pos

        pos_rate = pos / total_pos if total_pos > 0 else 0
        neg_rate = neg / total_neg if total_neg > 0 else 0

        # 避免 ln(0)，加平滑
        pos_rate = max(pos_rate, 0.0001)
        neg_rate = max(neg_rate, 0.0001)

        woe = np.log(pos_rate / neg_rate)
        iv_bin = (pos_rate - neg_rate) * woe
        iv_total += iv_bin

        bins_info.append({
            'bin': str(name),
            'count': len(group),
            'pos': pos,
            'neg': neg,
            'pos_rate': pos / len(group) if len(group) > 0 else 0,
            'woe': woe,
        })

    return {
        'iv': iv_total,
        'coverage': len(s) / len(series),
        'n_bins': len(bins_info),
        'bins_info': bins_info,
    }


def calc_iv_all_features(X, y, n_bins=5):
    """计算所有特征的 IV 值"""
    results = []
    for col in X.columns:
        info = calc_iv_for_feature(X[col], y, n_bins)
        results.append({
            'feature': col,
            'iv': info.get('iv', 0),
            'coverage': info.get('coverage', 0),
            'n_bins': info.get('n_bins', 0),
            'skip': info.get('skip', False),
        })

    iv_df = pd.DataFrame(results).sort_values('iv', ascending=False).reset_index(drop=True)
    return iv_df


# ====== Step 2: LightGBM 重要性交叉 ======

def cross_reference_lgbm(iv_df, output_dir):
    """用 LightGBM 重要性交叉验证 IV 排名"""
    importance_path = os.path.join(output_dir, 'lgb_feature_importance_full.csv')
    if not os.path.exists(importance_path):
        print('  ⚠️ 无 LightGBM 重要性文件，跳过交叉验证')
        iv_df['lgb_rank'] = 999
        iv_df['combined_rank'] = iv_df.index + 1
        return iv_df

    lgb_imp = pd.read_csv(importance_path)
    imp_dict = dict(zip(lgb_imp['feature'], lgb_imp['importance']))

    iv_df['lgb_importance'] = iv_df['feature'].map(imp_dict).fillna(0)
    iv_df['iv_rank'] = range(1, len(iv_df) + 1)
    iv_df['lgb_rank'] = iv_df['lgb_importance'].rank(ascending=False).astype(int)
    iv_df['combined_rank'] = (iv_df['iv_rank'] + iv_df['lgb_rank']) / 2
    iv_df = iv_df.sort_values('combined_rank').reset_index(drop=True)
    return iv_df


# ====== Step 3: 去相关 ======

def remove_correlated(X, features, iv_df, threshold=0.7):
    """贪心去相关: 按 IV 降序，依次保留与已选特征 |r| < threshold 的"""
    iv_order = iv_df[iv_df['feature'].isin(features)].set_index('feature')['iv']
    sorted_features = iv_order.sort_values(ascending=False).index.tolist()

    selected = []
    for feat in sorted_features:
        if feat not in X.columns:
            continue
        # 检查与已选特征的相关性
        is_correlated = False
        for existing in selected:
            valid = X[feat].notna() & X[existing].notna()
            if valid.sum() < 30:
                continue
            corr = X.loc[valid, feat].corr(X.loc[valid, existing])
            if abs(corr) > threshold:
                is_correlated = True
                break
        if not is_correlated:
            selected.append(feat)

    return selected


# ====== Step 3.5: VIF 筛选 ======

def filter_by_vif(X, features, max_vif=5.0):
    """逐步回归式 VIF 筛选: 每次剔除 VIF 最大的特征，直到所有 VIF < max_vif

    Args:
        X: DataFrame
        features: 候选特征列表
        max_vif: VIF 阈值，默认5.0（严格标准，10.0为宽松标准）

    Returns:
        筛选后的特征列表
    """
    from statsmodels.stats.outliers_influence import variance_inflation_factor

    remaining = list(features)

    for _ in range(len(features)):
        X_sel = X[remaining].copy()
        X_sel = X_sel.replace([np.inf, -np.inf], np.nan).fillna(X_sel.median())

        # 计算 VIF
        vif_values = []
        for i, col in enumerate(remaining):
            try:
                vif = variance_inflation_factor(X_sel.values, i)
                vif_values.append(vif if np.isfinite(vif) else 999)
            except Exception:
                vif_values.append(999)

        max_vif_val = max(vif_values)
        if max_vif_val <= max_vif:
            break

        # 剔除 VIF 最大的特征
        drop_idx = vif_values.index(max_vif_val)
        drop_feat = remaining[drop_idx]
        print(f'    VIF剔除: {drop_feat:25s} VIF={max_vif_val:.1f}')
        remaining.pop(drop_idx)

    # 输出最终 VIF
    print(f'    VIF筛选后: {len(remaining)}个特征 (阈值<{max_vif})')
    X_final = X[remaining].replace([np.inf, -np.inf], np.nan).fillna(X[remaining].median())
    for i, col in enumerate(remaining):
        try:
            vif = variance_inflation_factor(X_final.values, i)
            print(f'      {col:25s} VIF={vif:.2f}')
        except Exception:
            pass

    return remaining


# ====== Step 5: WOE 变换 ======

def woe_transform(X, y, features, n_bins=5):
    """对选定特征进行 WOE 变换

    Returns:
        X_woe: WOE 变换后的 DataFrame
        bins_dict: {feature: {bins: [...], woe_map: {bin_str: woe_value}, edges: [...]}}
    """
    X_woe = X.copy()
    bins_dict = {}

    for feat in features:
        valid = X[feat].notna() & y.notna()
        s = X.loc[valid, feat]
        y_valid = y.loc[valid]

        if len(s) < 50:
            continue

        try:
            bin_series = pd.qcut(s, n_bins, duplicates='drop')
        except ValueError:
            continue

        df_bins = pd.DataFrame({'bin': bin_series, 'y': y_valid.values})
        total_pos = df_bins['y'].sum()
        total_neg = len(df_bins) - total_pos

        woe_map = {}
        bin_edges = []
        for name, group in df_bins.groupby('bin', observed=True):
            pos = group['y'].sum()
            neg = len(group) - pos
            pos_rate = max(pos / total_pos, 0.0001) if total_pos > 0 else 0.0001
            neg_rate = max(neg / total_neg, 0.0001) if total_neg > 0 else 0.0001
            woe = np.log(pos_rate / neg_rate)
            woe_map[str(name)] = woe
            # 提取分箱边界
            if hasattr(name, 'left'):
                bin_edges.append((name.left, name.right))

        bins_dict[feat] = {
            'woe_map': woe_map,
            'edges': sorted(bin_edges, key=lambda x: x[0]) if bin_edges else [],
        }

        # 变换：用相同分箱对全量数据(含 NaN)做 WOE 替换
        full_bins = pd.qcut(X[feat], n_bins, duplicates='drop')
        X_woe[feat] = full_bins.astype(str).map(woe_map)
        # 原始 NaN 保持 NaN

    return X_woe, bins_dict


# ====== Step 6: 训练评分卡 LR ======

def train_scorecard_lr(X_woe, y, features):
    """训练评分卡逻辑回归"""
    from sklearn.linear_model import LogisticRegression
    from sklearn.model_selection import StratifiedKFold, cross_val_score, cross_val_predict
    from sklearn.metrics import classification_report, roc_auc_score

    X_sel = X_woe[features].copy()
    X_sel = X_sel.replace([np.inf, -np.inf], np.nan).fillna(0)

    model = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)

    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    auc_scores = cross_val_score(model, X_sel, y, cv=cv, scoring='roc_auc')
    auc_mean = auc_scores.mean()
    auc_std = auc_scores.std()

    print(f'  5折CV AUC: {auc_mean:.3f} ± {auc_std:.3f}')

    y_pred = cross_val_predict(model, X_sel, y, cv=cv)
    print(f'  分类报告:')
    print(classification_report(y, y_pred, target_names=['亏损', '盈利']))

    # 全量训练
    model.fit(X_sel, y)

    return model, auc_mean, auc_std


# ====== Step 7: 生成评分卡 ======

def build_scorecard(model, bins_dict, features, base_score=600, pdo=20):
    """生成评分卡得分表

    公式:
        B = pdo / ln(2)
        基础分 = base_score - B * intercept
        每特征每箱得分 = B * coef * woe
    """
    import math
    B = pdo / math.log(2)
    intercept = model.intercept_[0]
    coefs = dict(zip(features, model.coef_[0]))
    base_points = base_score - B * intercept

    rows = []
    for feat in features:
        coef = coefs.get(feat, 0)
        feat_info = bins_dict.get(feat, {})
        woe_map = feat_info.get('woe_map', {})

        for bin_str, woe in sorted(woe_map.items()):
            points = round(B * coef * woe, 1)
            rows.append({
                '特征': feat,
                '分箱范围': bin_str,
                'WOE': round(woe, 4),
                '系数': round(coef, 4),
                '得分': points,
            })

    # 基础分行
    rows.append({
        '特征': '基础分',
        '分箱范围': '-',
        'WOE': '-',
        '系数': '-',
        '得分': round(base_points, 1),
    })

    scorecard_df = pd.DataFrame(rows)
    return scorecard_df, base_points, B


# ====== Step 8: 评估 ======

def evaluate_scorecard(X_woe, y, model, features, base_points, B):
    """评估评分卡得分分布"""
    from sklearn.metrics import roc_auc_score

    X_sel = X_woe[features].copy()
    X_sel = X_sel.replace([np.inf, -np.inf], np.nan).fillna(0)

    # 计算每条样本的总分
    coefs = dict(zip(features, model.coef_[0]))
    scores = np.full(len(X_sel), base_points)
    for feat in features:
        woe_vals = X_sel[feat].values
        scores += B * coefs[feat] * woe_vals
    scores = np.nan_to_num(scores, nan=base_points)

    # AUC
    valid = y.notna()
    auc = roc_auc_score(y[valid], scores[valid]) if valid.sum() > 50 else 0

    # KS
    scores_pos = scores[y == 1]
    scores_neg = scores[y == 0]
    try:
        from scipy.stats import ks_2samp
        ks_stat, ks_p = ks_2samp(scores_pos, scores_neg)
    except Exception:
        ks_stat = 0

    # 得分分布表
    bins_edges = [-np.inf, 550, 575, 600, 625, 650, np.inf]
    bins_labels = ['<550', '550-575', '575-600', '600-625', '625-650', '>650']
    score_groups = pd.cut(scores, bins=bins_edges, labels=bins_labels)

    dist_df = pd.DataFrame({
        '得分区间': score_groups,
        '实际': y.values,
    })
    dist_summary = dist_df.groupby('得分区间', observed=True).agg(
        样本数=('实际', 'count'),
        盈利数=('实际', 'sum'),
    ).reset_index()
    dist_summary['盈利率'] = (dist_summary['盈利数'] / dist_summary['样本数'] * 100).round(1).astype(str) + '%'

    print(f'\n  AUC (得分): {auc:.3f}')
    print(f'  KS 统计量: {ks_stat:.3f}')
    print(f'\n  得分分布:')
    print(dist_summary.to_string(index=False))

    # 盈利/亏损组平均分
    print(f'\n  盈利组平均分: {scores[y == 1].mean():.1f}')
    print(f'  亏损组平均分: {scores[y == 0].mean():.1f}')

    return {
        'auc': auc,
        'ks': ks_stat,
        'score_dist': dist_summary,
        'scores': scores,
        'mean_score_profit': scores[y == 1].mean(),
        'mean_score_loss': scores[y == 0].mean(),
    }


def save_scorecard_artifacts(output_dir, model, bins_dict, features, medians,
                             scorecard_df, eval_results, base_points, B, args, iv_df):
    """保存所有输出文件"""
    # 1. 评分卡模型 pkl
    pkl_path = os.path.join(output_dir, 'scorecard_model.pkl')
    with open(pkl_path, 'wb') as f:
        pickle.dump({
            'model': model,
            'features': features,
            'woe_bins': bins_dict,
            'medians': medians,
            'scoring_params': {
                'base_score': 600,
                'pdo': 20,
                'B': B,
                'base_points': base_points,
            },
        }, f)
    print(f'  模型: {pkl_path}')

    # 2. WOE 分箱定义 JSON
    bins_path = os.path.join(output_dir, 'scorecard_bins.json')
    # 转换 edges 中的 float 为 string 以便 JSON 序列化
    bins_serializable = {}
    for feat, info in bins_dict.items():
        bins_serializable[feat] = {
            'woe_map': {str(k): float(v) for k, v in info['woe_map'].items()},
            'edges': [[float(e[0]), float(e[1])] for e in info.get('edges', [])],
        }
    with open(bins_path, 'w', encoding='utf-8') as f:
        json.dump(bins_serializable, f, ensure_ascii=False, indent=2)
    print(f'  分箱: {bins_path}')

    # 3. 评分卡得分表 CSV
    sc_path = os.path.join(output_dir, 'scorecard_table.csv')
    scorecard_df.to_csv(sc_path, index=False, encoding='utf-8-sig')
    print(f'  评分卡: {sc_path}')

    # 4. IV 分析 CSV
    iv_path = os.path.join(output_dir, 'iv_analysis.csv')
    iv_df.to_csv(iv_path, index=False, encoding='utf-8-sig')
    print(f'  IV分析: {iv_path}')

    # 5. 评分卡报告
    report_path = os.path.join(output_dir, 'scorecard_report.txt')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write('定增盈利预测 — 评分卡模型报告\n')
        f.write('=' * 60 + '\n\n')
        f.write(f'训练时间: {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")}\n')
        f.write(f'输入: {args.features_path}\n')
        f.write(f'盈利阈值: {args.threshold}%\n')
        f.write(f'样本: {eval_results.get("n_samples", "?")}条')
        f.write(f' (盈利{eval_results.get("n_profit", "?")} / 亏损{eval_results.get("n_loss", "?")})\n\n')

        f.write('特征选择过程:\n')
        f.write(f'  Step 1: IV筛选 (>= {args.iv_min}) → {eval_results.get("n_after_iv", "?")}个特征\n')
        f.write(f'  Step 2: LightGBM重要性交叉验证\n')
        f.write(f'  Step 3: 去相关 (|r| > 0.7) → {eval_results.get("n_after_corr", "?")}个特征\n')
        f.write(f'  Step 4: 取Top {len(features)}个\n\n')

        f.write(f'最终 {len(features)} 个特征:\n')
        for i, feat in enumerate(features, 1):
            iv_val = iv_df.loc[iv_df['feature'] == feat, 'iv'].values
            iv_str = f'{iv_val[0]:.4f}' if len(iv_val) > 0 else '?'
            f.write(f'  {i:2d}. {feat:25s} IV={iv_str}\n')

        f.write(f'\n模型性能:\n')
        f.write(f'  5折CV AUC: {eval_results.get("auc_cv", "?"):.3f}\n')
        f.write(f'  得分AUC: {eval_results.get("auc", "?"):.3f}\n')
        f.write(f'  KS统计量: {eval_results.get("ks", "?"):.3f}\n')
        f.write(f'  盈利组平均分: {eval_results.get("mean_score_profit", "?"):.1f}\n')
        f.write(f'  亏损组平均分: {eval_results.get("mean_score_loss", "?"):.1f}\n\n')

        f.write(f'得分分布:\n')
        dist = eval_results.get('score_dist')
        if dist is not None:
            f.write(dist.to_string(index=False) + '\n\n')

        f.write(f'评分卡得分表见: scorecard_table.csv\n')
        f.write(f'IV分析见: iv_analysis.csv\n')
    print(f'  报告: {report_path}')

    # 6. 版本归档 + 注册到 model_registry
    _archive_and_register_scorecard(
        output_dir, features, eval_results, args,
        files=[pkl_path, bins_path, sc_path, iv_path, report_path],
    )


def _archive_and_register_scorecard(output_dir, features, eval_results, args, files):
    """把评分卡产物归档到版本子目录，并注册到 model_registry。"""
    import shutil
    from datetime import datetime

    auc = eval_results.get('auc', 0) or 0
    date_str = datetime.now().strftime('%Y%m%d_%H%M')
    auc_tag = f'auc{auc:.2f}'.replace('.', '')
    version_name = f'v_{date_str}_scorecard_{len(features)}feat_{auc_tag}'
    version_dir = os.path.join(output_dir, version_name)
    os.makedirs(version_dir, exist_ok=True)

    archived = []
    for fpath in files:
        if os.path.exists(fpath):
            shutil.copy2(fpath, os.path.join(version_dir, os.path.basename(fpath)))
            archived.append(os.path.basename(fpath))

    # VERSION.md
    with open(os.path.join(version_dir, 'VERSION.md'), 'w', encoding='utf-8') as f:
        f.write(f'# 评分卡版本: {version_name}\n\n')
        f.write(f'**训练时间**: {datetime.now().strftime("%Y-%m-%d %H:%M")}\n')
        f.write(f'**入模特征数**: {len(features)}\n')
        f.write(f'**训练样本**: {eval_results.get("n_samples", "?")} '
                f'(盈利{eval_results.get("n_profit", "?")} / 亏损{eval_results.get("n_loss", "?")})\n')
        f.write(f'**盈利阈值**: {args.threshold}%\n')
        f.write(f'**5折CV AUC**: {eval_results.get("auc_cv", "?"):.3f}\n')
        f.write(f'**得分AUC**: {eval_results.get("auc", "?"):.3f}\n')
        f.write(f'**KS**: {eval_results.get("ks", "?"):.3f}\n\n')
        f.write('## 入模特征\n')
        for x in features:
            f.write(f'- {x}\n')

    try:
        from deploy.model_registry import register_version
        n = eval_results.get('n_samples')
        pos = eval_results.get('n_profit')
        pos_rate = float(pos) / float(n) if n else None
        register_version(
            model_type='scorecard',
            version=version_name,
            dir=version_dir,
            metrics={'auc_cv': float(eval_results.get('auc_cv', 0) or 0),
                     'auc': float(eval_results.get('auc', 0) or 0),
                     'ks': float(eval_results.get('ks', 0) or 0)},
            n_features=len(features),
            threshold=args.threshold,
            n_samples=n,
            positive_rate=pos_rate,
            files=archived,
            note=f'评分卡 {len(features)}特征',
            set_current=True,
        )
        print(f'  版本归档: {version_dir}')
        print(f'  registry: 已注册为 scorecard 当前版本 → {version_name}')
    except Exception as e:
        print(f'  ⚠ registry 注册失败(不影响模型文件): {e}')


# ====== 主流程 ======

def main():
    parser = argparse.ArgumentParser(description='定增盈利预测 - 评分卡模型训练')
    parser.add_argument('features_path', help='features_derived.parquet 路径')
    parser.add_argument('--threshold', type=float, default=-10, help='盈利阈值(%%)，默认-10')
    parser.add_argument('--n-features', type=int, default=12, help='最终特征数，默认12')
    parser.add_argument('--iv-min', type=float, default=0.05, help='IV最低阈值，默认0.05')
    args = parser.parse_args()

    output_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'output')
    os.makedirs(output_dir, exist_ok=True)

    # ── 加载数据 ──
    from train.train_models import prepare_features_full
    print('1. 加载特征数据...')
    if args.features_path.endswith('.parquet'):
        df = pd.read_parquet(args.features_path)
    else:
        df = pd.read_csv(args.features_path)
    print(f'   {len(df)} 条记录, {len(df.columns)} 列')

    X, y = prepare_features_full(df, args.threshold)
    if X is None:
        sys.exit(1)

    # 统一排除清单 (与 compare_selection.py 共用 feature_exclusions)
    from features.feature_exclusions import get_excluded_columns
    excl = get_excluded_columns(X.columns)
    if excl:
        X = X.drop(columns=excl)
        print(f'   排除泄漏/artifact/业务字段: {excl}')

    n_total_features = len(X.columns)

    # ── Step 1: IV 计算 ──
    print(f'\n2. IV计算 ({n_total_features}个特征)...')
    iv_df = calc_iv_all_features(X, y)
    n_after_iv = (iv_df['iv'] >= args.iv_min).sum()
    print(f'   IV >= {args.iv_min}: {n_after_iv}个特征')
    # 显示 Top 15
    print(f'   IV Top 15:')
    for _, row in iv_df.head(15).iterrows():
        print(f'     {row["feature"]:25s} IV={row["iv"]:.4f}')

    # ── Step 2: LightGBM 交叉验证 ──
    print(f'\n3. LightGBM重要性交叉验证...')
    iv_df = cross_reference_lgbm(iv_df, output_dir)

    # ── Step 3: IV 筛选 + 去相关 + VIF ──
    print(f'\n4. 特征筛选: IV>={args.iv_min} → 去相关 → VIF → Top {args.n_features}...')
    iv_pass = iv_df[iv_df['iv'] >= args.iv_min]['feature'].tolist()
    print(f'   IV筛选后: {len(iv_pass)}个')

    decorrelated = remove_correlated(X, iv_pass, iv_df, threshold=0.7)
    print(f'   去相关后: {len(decorrelated)}个')

    # VIF 筛选（排除多重共线性）
    print(f'   VIF筛选 (阈值<5.0)...')
    vif_filtered = filter_by_vif(X, decorrelated, max_vif=5.0)

    final_features = vif_filtered[:args.n_features]
    print(f'   最终选择: {len(final_features)}个')
    for i, f in enumerate(final_features, 1):
        iv_val = iv_df.loc[iv_df['feature'] == f, 'iv'].values[0]
        print(f'     {i:2d}. {f:25s} IV={iv_val:.4f}')

    # ── Step 5: WOE 变换 ──
    print(f'\n5. WOE变换...')
    X_woe, bins_dict = woe_transform(X, y, final_features)

    # ── Step 6: 训练评分卡 LR ──
    print(f'\n6. 训练评分卡逻辑回归 ({len(final_features)}个WOE特征)...')
    model, auc_cv, auc_std = train_scorecard_lr(X_woe, y, final_features)

    # ── Step 7: 生成评分卡 ──
    print(f'\n7. 生成评分卡得分表 (Base=600, PDO=20)...')
    scorecard_df, base_points, B = build_scorecard(model, bins_dict, final_features)
    print(scorecard_df.to_string(index=False))

    # ── Step 8: 评估 ──
    print(f'\n8. 评估评分卡...')
    eval_results = evaluate_scorecard(X_woe, y, model, final_features, base_points, B)

    # 补充评估结果
    eval_results.update({
        'n_samples': len(y),
        'n_profit': int(y.sum()),
        'n_loss': int((1 - y).sum()),
        'n_after_iv': n_after_iv,
        'n_after_corr': len(decorrelated),
        'auc_cv': auc_cv,
    })

    # ── 保存 ──
    print(f'\n9. 保存输出...')
    medians = {c: float(X[c].median()) for c in final_features}
    save_scorecard_artifacts(output_dir, model, bins_dict, final_features, medians,
                             scorecard_df, eval_results, base_points, B, args, iv_df)

    print(f'\n✅ 评分卡训练完成!')


if __name__ == '__main__':
    main()
