#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增盈利预测模型训练与评估

基于定增筛选结果(含PE估值、行业对比、财务评分等指标)，
训练LightGBM模型预测7个月后涨跌幅/盈利概率。

用法:
    python ml_training/train_profitability_model.py <scored_excel_path> [--threshold -10]

输出:
    - 控制台: 交叉验证AUC、分类报告、特征重要性
    - ml_training/output/: 模型文件、特征重要性图、预测结果
"""

import sys
import os
import argparse
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')

# 特征定义
NUMERIC_FEATURES = [
    # PE估值
    '溢价率下限', '溢价率上限', '有效阈值数',
    # 财务评分（取最新年）
    '总分_最新', '盈利能力_最新', '成长能力_最新',
    '总分_斜率', '盈利能力_斜率', '成长能力_斜率',
    # 报价日价格
    '报价日价格',
]

CATEGORICAL_FEATURES = [
    '一级行业',
]

SUB_SCENARIO_FEATURES = [
    '市场指数', '行业PE', '个股PE',
    'DCF估值', '修正PE估值',
    '参数构造', '蒙特卡洛', '反向推算',
]


def load_and_prepare(filepath, threshold=-10):
    """加载scored Excel，构建特征和标签。

    Args:
        filepath: scored Excel路径
        threshold: 盈利阈值(%)，涨跌幅>threshold为盈利
    """
    df = pd.read_excel(filepath, sheet_name='Sheet1')
    print(f'读取 {len(df)} 条记录')

    # ===== 构建标签 =====
    ret_col = '7个月后涨跌幅'
    if ret_col not in df.columns:
        print(f'❌ 找不到 {ret_col} 列')
        return None, None, None

    # 转数值
    df[ret_col] = pd.to_numeric(
        df[ret_col].astype(str).str.replace('%', '').str.replace('+', ''),
        errors='coerce'
    )
    df = df.dropna(subset=[ret_col]).reset_index(drop=True)
    print(f'有效涨跌幅数据: {len(df)} 条')

    y_reg = df[ret_col]  # 回归目标: 涨跌幅
    y_cls = (df[ret_col] * 100 > threshold).astype(int)  # 分类目标: 盈利=1
    print(f'盈利占比(threshold={threshold}%): {y_cls.mean()*100:.1f}%')

    # ===== 构建特征 =====
    features = pd.DataFrame(index=df.index)

    # 1. 数值特征
    # 溢价率（转数值）
    for col in ['溢价率下限', '溢价率上限']:
        if col in df.columns:
            features[col] = pd.to_numeric(
                df[col].astype(str).str.replace('%', '').str.replace('+', ''),
                errors='coerce'
            )

    if '有效阈值数' in df.columns:
        features['有效阈值数'] = pd.to_numeric(df['有效阈值数'], errors='coerce').fillna(0)

    # 报价日价格
    if '报价日价格' in df.columns:
        features['报价日价格'] = pd.to_numeric(df['报价日价格'], errors='coerce')

    # 2. 财务评分（找最新年份的列）
    score_cols = [c for c in df.columns if c.startswith('总分_') and c[3:].isdigit()]
    if score_cols:
        latest = max(c.split('_')[1] for c in score_cols)
        for prefix in ['总分', '盈利能力', '成长能力']:
            col = f'{prefix}_{latest}'
            if col in df.columns:
                features[f'{prefix}_最新'] = pd.to_numeric(df[col], errors='coerce').fillna(50)

    # 评分斜率
    for prefix in ['总分', '盈利能力', '成长能力']:
        col = f'{prefix}_斜率'
        if col in df.columns:
            features[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # 3. 子场景通过状态（✓=1, ✗=0）
    for col in SUB_SCENARIO_FEATURES:
        if col in df.columns:
            features[col] = df[col].astype(str).str.contains('✓').astype(int)

    # 子场景通过总数
    sub_pass_cols = [c for c in SUB_SCENARIO_FEATURES if c in features.columns]
    if sub_pass_cols:
        features['子场景通过数'] = features[sub_pass_cols].sum(axis=1)

    # 4. 类别特征（一级行业）
    if '一级行业' in df.columns:
        features['一级行业'] = df['一级行业'].fillna('未知').astype(str)

    # 5. 定增决策
    if '定增决策' in df.columns:
        features['定增建议参与'] = df['定增决策'].astype(str).str.contains('建议参与').astype(int)

    # 填充NaN
    for col in features.columns:
        if features[col].dtype in [np.float64, np.int64]:
            features[col] = features[col].fillna(features[col].median())

    print(f'\n特征数: {len(features.columns)}')
    print(f'特征列: {list(features.columns)}')

    return features, y_cls, y_reg


def train_classification(X, y, output_dir):
    """训练分类模型（预测盈利/亏损）。"""
    from sklearn.model_selection import StratifiedKFold, cross_val_score, cross_val_predict
    from sklearn.metrics import classification_report, roc_auc_score, confusion_matrix
    import lightgbm as lgb

    print('\n' + '='*60)
    print('分类模型：预测盈利概率')
    print('='*60)

    # 类别特征编码
    cat_cols = [c for c in ['一级行业'] if c in X.columns]
    for c in cat_cols:
        X[c] = X[c].astype('category')

    model = lgb.LGBMClassifier(
        n_estimators=200, max_depth=5, learning_rate=0.05,
        num_leaves=31, subsample=0.8, colsample_bytree=0.8,
        reg_alpha=0.1, reg_lambda=0.1, random_state=42, verbose=-1
    )

    # 交叉验证
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    auc_scores = cross_val_score(model, X, y, cv=cv, scoring='roc_auc')
    print(f'\n5折交叉验证 AUC: {auc_scores.mean():.3f} ± {auc_scores.std():.3f}')

    # 详细预测
    y_pred = cross_val_predict(model, X, y, cv=cv)
    y_proba = cross_val_predict(model, X, y, cv=cv, method='predict_proba')[:, 1]

    print(f'\n分类报告:')
    print(classification_report(y, y_pred, target_names=['亏损', '盈利']))

    cm = confusion_matrix(y, y_pred)
    print(f'混淆矩阵:')
    print(f'  预测亏损  预测盈利')
    print(f'实际亏损  {cm[0,0]:6d}  {cm[0,1]:6d}')
    print(f'实际盈利  {cm[1,0]:6d}  {cm[1,1]:6d}')

    # 全量训练
    model.fit(X, y)

    # 特征重要性
    importance = pd.DataFrame({
        'feature': X.columns,
        'importance': model.feature_importances_
    }).sort_values('importance', ascending=False)

    print(f'\n特征重要性 Top 15:')
    print(importance.head(15).to_string(index=False))

    # 保存模型
    model.booster_.save_model(os.path.join(output_dir, 'lgb_classifier.txt'))

    # 保存特征重要性图
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        fig, ax = plt.subplots(figsize=(10, 8))
        lgb.plot_importance(model, max_num_features=20, ax=ax)
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, 'feature_importance.png'), dpi=150)
        plt.close()
        print(f'特征重要性图已保存')
    except Exception:
        pass

    return model, importance


def train_regression(X, y_reg, output_dir):
    """训练回归模型（预测涨跌幅）。"""
    from sklearn.model_selection import KFold, cross_val_score
    import lightgbm as lgb

    print('\n' + '='*60)
    print('回归模型：预测7个月涨跌幅')
    print('='*60)

    cat_cols = [c for c in ['一级行业'] if c in X.columns]
    for c in cat_cols:
        X[c] = X[c].astype('category')

    model = lgb.LGBMRegressor(
        n_estimators=300, max_depth=5, learning_rate=0.03,
        num_leaves=31, subsample=0.8, colsample_bytree=0.8,
        reg_alpha=0.1, reg_lambda=0.5, random_state=42, verbose=-1
    )

    cv = KFold(n_splits=5, shuffle=True, random_state=42)
    r2_scores = cross_val_score(model, X, y_reg, cv=cv, scoring='r2')
    mae_scores = cross_val_score(model, X, y_reg, cv=cv, scoring='neg_mean_absolute_error')

    print(f'\n5折交叉验证:')
    print(f'  R²: {r2_scores.mean():.3f} ± {r2_scores.std():.3f}')
    print(f'  MAE: {-mae_scores.mean():.4f} (平均绝对误差)')

    model.fit(X, y_reg)
    model.booster_.save_model(os.path.join(output_dir, 'lgb_regressor.txt'))

    return model


def analyze_by_group(X, y_cls, y_reg, output_dir):
    """按子场景/行业分组分析盈利占比。"""
    print('\n' + '='*60)
    print('分组分析：各子场景的盈利预测力')
    print('='*60)

    df_analysis = X.copy()
    df_analysis['盈利'] = y_cls.values
    df_analysis['涨跌幅'] = y_reg.values

    # 子场景分析
    sub_cols = [c for c in SUB_SCENARIO_FEATURES if c in df_analysis.columns]
    for col in sub_cols:
        for val, label in [(1, '通过'), (0, '不通过')]:
            sub = df_analysis[df_analysis[col] == val]
            if len(sub) > 0:
                win_rate = sub['盈利'].mean() * 100
                avg_ret = sub['涨跌幅'].mean()
                print(f'  {col:8s} {label}: 盈利占比={win_rate:5.1f}%  平均收益={avg_ret:+.4f}  n={len(sub)}')

    # 行业分析
    if '一级行业' in df_analysis.columns:
        print(f'\n行业盈利占比 Top 10:')
        ind_stats = df_analysis.groupby('一级行业').agg(
            总数=('盈利', 'count'),
            盈利数=('盈利', 'sum'),
            平均涨跌=('涨跌幅', 'mean'),
        )
        ind_stats['盈利占比'] = (ind_stats['盈利数'] / ind_stats['总数'] * 100).round(1)
        ind_stats = ind_stats.sort_values('总数', ascending=False).head(10)
        print(ind_stats.to_string())

        # 保存
        ind_stats.to_csv(os.path.join(output_dir, 'industry_analysis.csv'))


def main():
    parser = argparse.ArgumentParser(description='定增盈利预测模型训练')
    parser.add_argument('excel_path', help='scored Excel路径')
    parser.add_argument('--threshold', type=float, default=-10,
                        help='盈利阈值(%%)，涨跌幅>此值为盈利，默认-10')
    args = parser.parse_args()

    # 输出目录
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    os.makedirs(output_dir, exist_ok=True)

    # 加载数据
    X, y_cls, y_reg = load_and_prepare(args.excel_path, args.threshold)
    if X is None:
        sys.exit(1)

    # 训练分类模型
    cls_model, importance = train_classification(X.copy(), y_cls, output_dir)

    # 训练回归模型
    reg_model = train_regression(X.copy(), y_reg, output_dir)

    # 分组分析
    analyze_by_group(X, y_cls, y_reg, output_dir)

    # 保存特征重要性
    importance.to_csv(os.path.join(output_dir, 'feature_importance.csv'), index=False)

    print(f'\n✅ 完成! 输出目录: {output_dir}')
    print(f'   - lgb_classifier.txt (分类模型)')
    print(f'   - lgb_regressor.txt (回归模型)')
    print(f'   - feature_importance.png (特征重要性图)')
    print(f'   - feature_importance.csv (特征重要性数据)')
    print(f'   - industry_analysis.csv (行业分析)')


if __name__ == '__main__':
    main()
