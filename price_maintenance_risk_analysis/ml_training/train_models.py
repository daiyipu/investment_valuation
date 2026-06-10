#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增盈利预测 - 双模型训练(LightGBM + 逻辑回归)

LightGBM: 精准度高，自动捕捉非线性交互
逻辑回归: 可读性强，系数表类似评分卡

用法:
    python ml_training/train_models.py <features.parquet> [--threshold -10]

输出:
    ml_training/output/
    ├── lgb_classifier.txt      # LightGBM模型
    ├── lr_classifier.pkl       # 逻辑回归模型
    ├── lgb_feature_importance.png
    ├── lr_coefficients.csv     # 逻辑回归系数表(评分卡)
    ├── model_comparison.png    # 两模型对比图
    └── evaluation_report.txt   # 评估报告
"""

import sys
import os
import argparse
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')


# ====== 特征定义 ======
# 核心特征（去除高基数噪声：报价日价格/溢价率等）
FEATURE_COLUMNS = [
    # 财务趋势（最重要）
    '总分_斜率', '盈利能力_斜率', '成长能力_斜率',
    # 财务绝对值
    '总分_T', '盈利能力_T', '成长能力_T',
    # 行情特征
    '波动率_60d', '波动率_120d', '波动率_250d',
    '年化收益_60d', '年化收益_120d',
    '区间收益_60d', '区间收益_120d',
    '胜率_60d', '胜率_120d', '胜率_250d',
    '漂移率', '波动率', '换手率', '数据天数',
    # 估值特征
    '个股PE', '个股PB', '个股PS',
    '行业PE', '行业PB', '行业PS',
    # FCF特征
    '营收_T', 'NOPAT_T', 'FCF_T',
    '营收_T1', 'NOPAT_T1', 'FCF_T1',
    # 定增决策
    '有效阈值数', 'step1通过', 'step2通过', 'step3通过',
    # 子场景
    '市场指数_通过', '行业PE_通过', '个股PE_通过',
    'DCF估值_通过', '修正PE估值_通过',
    '参数构造_通过', '蒙特卡洛_通过', '反向推算_通过',
    # 定增参数
    '锁定期', '净债务', '净利润',
]

CATEGORICAL_COLS = ['一级行业']

# MySQL财务比率（如果有的话）
MYSQL_RATIO_COLS = [
    '流动比率', '速动比率', '存货周转率', '应收账款周转率',
    '总资产周转率', 'ROA', 'ROE', 'ROE摊薄',
    '净利率', '毛利率', '资产负债率', '产权比率',
    '已获利息倍数', '研发费用率',
    '营收增长', '净利增长', '扣非净利增长', 'ROE增长',
]


def load_features(filepath):
    """加载特征数据"""
    if filepath.endswith('.parquet'):
        df = pd.read_parquet(filepath)
    else:
        df = pd.read_csv(filepath)
    print(f'加载 {len(df)} 条记录, {len(df.columns)} 列')
    return df


def prepare_features(df, threshold=-10):
    """准备特征矩阵X和标签y"""
    # 标签
    label_col = f'标签_盈利_{threshold}'
    if label_col not in df.columns:
        # 尝试从涨跌幅计算
        if '7个月涨跌幅' in df.columns:
            df[label_col] = (df['7个月涨跌幅'] > threshold / 100).astype(int)
        else:
            print(f'❌ 找不到标签列 {label_col}')
            return None, None

    y = df[label_col]
    valid = y.notna()
    df = df[valid].reset_index(drop=True)
    y = y[valid].reset_index(drop=True)

    # 选择可用特征
    available = []
    for c in FEATURE_COLUMNS + MYSQL_RATIO_COLS:
        if c in df.columns:
            available.append(c)

    # 类别特征
    cat_available = [c for c in CATEGORICAL_COLS if c in df.columns]

    X = df[available + cat_available].copy()

    # 数值列填充
    for c in available:
        X[c] = pd.to_numeric(X[c], errors='coerce')
        X[c] = X[c].fillna(X[c].median())

    # 类别列
    for c in cat_available:
        X[c] = X[c].fillna('未知').astype(str)

    print(f'特征数: {len(available) + len(cat_available)} (数值{len(available)} + 类别{len(cat_available)})')
    print(f'盈利占比: {y.mean()*100:.1f}%')
    print(f'样本数: {len(y)} (盈利{y.sum():.0f} + 亏损{(1-y).sum():.0f})')

    if y.sum() < 5 or (1 - y).sum() < 5:
        print(f'⚠️ 样本太少或单一类别(盈利{y.sum():.0f}/亏损{(1-y).sum():.0f})，无法训练')
        print(f'   需要至少5个盈利+5个亏损样本。请用更多数据重跑。')
        return None, None

    return X, y


def train_lightgbm(X, y, output_dir):
    """LightGBM分类模型"""
    from sklearn.model_selection import StratifiedKFold, cross_val_score, cross_val_predict
    from sklearn.metrics import classification_report, roc_auc_score
    import lightgbm as lgb

    print('\n' + '='*60)
    print('LightGBM 模型（精准度高）')
    print('='*60)

    X_lgb = X.copy()
    for c in CATEGORICAL_COLS:
        if c in X_lgb.columns:
            X_lgb[c] = X_lgb[c].astype('category')

    model = lgb.LGBMClassifier(
        n_estimators=200, max_depth=5, learning_rate=0.05,
        num_leaves=31, subsample=0.8, colsample_bytree=0.8,
        reg_alpha=0.1, reg_lambda=0.1, random_state=42, verbose=-1
    )

    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    auc = cross_val_score(model, X_lgb, y, cv=cv, scoring='roc_auc')
    print(f'5折AUC: {auc.mean():.3f} ± {auc.std():.3f}')

    y_pred = cross_val_predict(model, X_lgb, y, cv=cv)
    print(f'\n分类报告:')
    print(classification_report(y, y_pred, target_names=['亏损', '盈利']))

    model.fit(X_lgb, y)
    model.booster_.save_model(os.path.join(output_dir, 'lgb_classifier.txt'))

    # 特征重要性
    importance = pd.DataFrame({
        'feature': X_lgb.columns, 'importance': model.feature_importances_
    }).sort_values('importance', ascending=False)
    print(f'\n特征重要性 Top 10:')
    print(importance.head(10).to_string(index=False))
    importance.to_csv(os.path.join(output_dir, 'lgb_feature_importance.csv'), index=False)

    # 画图
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        from matplotlib.font_manager import FontProperties
        fp = FontProperties(fname='/System/Library/Fonts/STHeiti Medium.ttc')
        plt.rcParams['font.family'] = fp.get_name()
        plt.rcParams['axes.unicode_minus'] = False
        fig, ax = plt.subplots(figsize=(10, 8))
        lgb.plot_importance(model, max_num_features=20, ax=ax)
        plt.title('LightGBM 特征重要性', fontsize=14)
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, 'lgb_feature_importance.png'), dpi=150)
        plt.close()
    except Exception:
        pass

    return model, auc.mean()


def train_logistic_regression(X, y, output_dir):
    """逻辑回归模型（可读性高，评分卡风格）"""
    from sklearn.model_selection import StratifiedKFold, cross_val_score, cross_val_predict
    from sklearn.metrics import classification_report
    from sklearn.linear_model import LogisticRegression
    from sklearn.preprocessing import StandardScaler
    import pickle

    print('\n' + '='*60)
    print('逻辑回归模型（可读性高，评分卡风格）')
    print('='*60)

    # 只用数值特征（逻辑回归不擅长类别特征）
    X_lr = X.copy()
    cat_cols = [c for c in CATEGORICAL_COLS if c in X_lr.columns]
    X_lr = X_lr.drop(columns=cat_cols)

    # 标准化
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X_lr)

    model = LogisticRegression(
        C=1.0, penalty='l2', max_iter=1000, random_state=42
    )

    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    auc = cross_val_score(model, X_scaled, y, cv=cv, scoring='roc_auc')
    print(f'5折AUC: {auc.mean():.3f} ± {auc.std():.3f}')

    y_pred = cross_val_predict(model, X_scaled, y, cv=cv)
    print(f'\n分类报告:')
    print(classification_report(y, y_pred, target_names=['亏损', '盈利']))

    model.fit(X_scaled, y)

    # 保存模型+scaler
    with open(os.path.join(output_dir, 'lr_classifier.pkl'), 'wb') as f:
        pickle.dump({'model': model, 'scaler': scaler, 'features': list(X_lr.columns)}, f)

    # 系数表（评分卡）
    coef_df = pd.DataFrame({
        '特征': X_lr.columns,
        '标准化系数': model.coef_[0],
    }).sort_values('标准化系数', ascending=False)

    # 转为"影响方向"和"影响强度"
    coef_df['影响'] = coef_df['标准化系数'].apply(
        lambda x: '↑正相关（高分→盈利）' if x > 0 else '↓负相关（高分→亏损）'
    )
    coef_df['影响强度'] = coef_df['标准化系数'].abs().apply(
        lambda x: '★★★' if x > 0.3 else ('★★' if x > 0.15 else '★')
    )
    coef_df = coef_df.rename(columns={'标准化系数': '系数(标准化后)'})

    print(f'\n逻辑回归系数表（评分卡）:')
    print(coef_df.head(15).to_string(index=False))
    coef_df.to_csv(os.path.join(output_dir, 'lr_coefficients.csv'), index=False)

    return model, scaler, auc.mean()


def main():
    parser = argparse.ArgumentParser(description='定增盈利预测 - 双模型训练')
    parser.add_argument('features_path', help='features.parquet 或 features.csv 路径')
    parser.add_argument('--threshold', type=float, default=-10, help='盈利阈值(%%)，默认-10')
    args = parser.parse_args()

    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    os.makedirs(output_dir, exist_ok=True)

    # 加载
    df = load_features(args.features_path)
    X, y = prepare_features(df, args.threshold)
    if X is None:
        sys.exit(1)

    # 训练LightGBM
    lgb_model, lgb_auc = train_lightgbm(X, y, output_dir)

    # 训练逻辑回归
    lr_model, lr_scaler, lr_auc = train_logistic_regression(X, y, output_dir)

    # 模型对比
    print('\n' + '='*60)
    print('模型对比')
    print('='*60)
    print(f'{"模型":<20} {"AUC":<10} {"优势"}')
    print(f'{"LightGBM":<20} {lgb_auc:<10.3f} 精准度高，捕捉非线性')
    print(f'{"逻辑回归":<20} {lr_auc:<10.3f} 可读性强，系数=评分卡')

    # 保存评估报告
    report_path = os.path.join(output_dir, 'evaluation_report.txt')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(f'定增盈利预测模型评估报告\n')
        f.write(f'盈利阈值: {args.threshold}%\n')
        f.write(f'样本数: {len(y)}, 盈利占比: {y.mean()*100:.1f}%\n\n')
        f.write(f'LightGBM AUC: {lgb_auc:.3f}\n')
        f.write(f'逻辑回归 AUC: {lr_auc:.3f}\n\n')
        f.write(f'逻辑回归系数表见: lr_coefficients.csv\n')
        f.write(f'LightGBM特征重要性见: lgb_feature_importance.csv\n')
    print(f'\n✅ 完成! 输出目录: {output_dir}')


if __name__ == '__main__':
    main()
