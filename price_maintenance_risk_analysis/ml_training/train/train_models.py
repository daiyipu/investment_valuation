#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增盈利预测 - 双模型训练(LightGBM + 逻辑回归)

LightGBM: 精准度高，自动捕捉非线性交互
逻辑回归: 可读性强，系数表类似评分卡

用法:
    python ml_training/train_models.py <features_derived.parquet> [--threshold -10]

输出:
    ml_training/output/
    ├── lgb_classifier.txt          # LightGBM模型
    ├── lgb_classifier_full.txt     # LightGBM模型（全量特征，predict_profitability使用）
    ├── lgb_full_meta.json          # 全量模型特征列表+median
    ├── lr_classifier.pkl           # 逻辑回归模型
    ├── lr_classifier_full.pkl      # 逻辑回归模型（全量特征，predict_profitability使用）
    ├── lgb_feature_importance.csv
    ├── lgb_feature_importance.png
    ├── lr_coefficients.csv         # 逻辑回归系数表(评分卡)
    └── evaluation_report.txt       # 评估报告
"""

import sys
import os
import json
import argparse
import warnings
import pickle
import shutil
from datetime import datetime
import numpy as np
import pandas as pd

# 与 scorecard/validate 共用唯一剔除清单(防多期限原始收益列泄漏)
PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # train/→ml_training/→PKG
for _p in (PKG, os.path.join(PKG,'ml_training'), os.path.join(PKG,'ml_training','pipeline'), os.path.join(PKG,'scripts')):
    if _p not in sys.path: sys.path.insert(0, _p)
from features.feature_exclusions import get_excluded_columns

warnings.filterwarnings('ignore')


# ====== 手工特征定义（旧版兼容，--classic 模式） ======
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

MYSQL_RATIO_COLS = [
    '流动比率', '速动比率', '存货周转率', '应收账款周转率',
    '总资产周转率', 'ROA', 'ROE', 'ROE摊薄',
    '净利率', '毛利率', '资产负债率', '产权比率',
    '已获利息倍数', '研发费用率',
    '营收增长', '净利增长', '扣非净利增长', 'ROE增长',
]

# 标签和目标列（训练时排除）
EXCLUDE_COLS = {'7个月涨跌幅', '7个月后价格', '标签_盈利_0', '标签_盈利_-10', '标签_盈利_-20',
                # 月线趋势: 个股 backward-looking 特征对该 forward-looking regime 标签无信号
                # (train IV≈0.03/0.00, 区分度反向); 测试验证期暂不入模, 留作 screen 用。非正式版, 不进 DROP_FIELDS。
                '月线MA10_slope3%', '月线趋势向上'}

# LightGBM 参数(降复杂度防过拟合): 原 300树/深5/叶31 在 train 1263×225 上 train AUC=1.0、OOT 仅 0.60 → 严重过拟合。
# 收紧: n_est 300→150, 深度 5→4, 叶 31→20, +min_child_samples=40, 配合早停。
LGB_PARAMS = dict(
    n_estimators=150, max_depth=4, num_leaves=20, learning_rate=0.03,
    min_child_samples=40, is_unbalance=True, subsample=0.8, colsample_bytree=0.8,
    reg_alpha=0.1, reg_lambda=0.1, random_state=42, verbose=-1,
)


def load_features(filepath):
    """加载特征数据"""
    if filepath.endswith('.parquet'):
        df = pd.read_parquet(filepath)
    else:
        df = pd.read_csv(filepath)
    print(f'加载 {len(df)} 条记录, {len(df.columns)} 列')
    return df


def prepare_features_full(df, threshold=-10, label_col=None):
    """全量数值特征（默认模式，与 validate_model.py 一致）。

    label_col: 显式标签列(如 标签_盈利_-10_3m / 标签_极性_灰度剔除_7m)；给定则直接用，
               None 时按 threshold 合成 标签_盈利_{threshold}。
    """
    if label_col is None:
        # 确保标签列名匹配（-10.0 → -10）
        threshold_int = int(threshold) if threshold == int(threshold) else threshold
        label_col = f'标签_盈利_{threshold_int}'
        if label_col not in df.columns and '7个月涨跌幅' in df.columns:
            df[label_col] = (df['7个月涨跌幅'] > threshold / 100).astype(int)
    if label_col not in df.columns:
        print(f'❌ 找不到标签列 {label_col}')
        return None, None

    y = df[label_col]
    valid = y.notna()
    df = df[valid].reset_index(drop=True)
    y = y[valid].reset_index(drop=True)

    # 全部数值列（排除标签/目标 + 统一剔除清单 feature_exclusions，防多期限原始收益列泄漏）
    exclude = set(get_excluded_columns(df.columns)) | EXCLUDE_COLS | {c for c in df.columns if '标签' in c}
    num_cols = [c for c in df.select_dtypes(include=[np.number]).columns if c not in exclude]

    X = df[num_cols].copy()
    for c in num_cols:
        X[c] = pd.to_numeric(X[c], errors='coerce').fillna(X[c].median())
    X = X.replace([np.inf, -np.inf], 0).fillna(0)

    print(f'特征数: {len(num_cols)} (全部数值特征)')
    print(f'盈利占比: {y.mean()*100:.1f}%')
    print(f'样本数: {len(y)} (盈利{y.sum():.0f} + 亏损{(1-y).sum():.0f})')

    if y.sum() < 5 or (1 - y).sum() < 5:
        print(f'⚠️ 样本太少或单一类别(盈利{y.sum():.0f}/亏损{(1-y).sum():.0f})，无法训练')
        return None, None

    return X, y


def prepare_features_classic(df, threshold=-10, label_col=None):
    """手工选择特征（--classic 模式，旧版兼容）"""
    if label_col is None:
        threshold_int = int(threshold) if threshold == int(threshold) else threshold
        label_col = f'标签_盈利_{threshold_int}'
        if label_col not in df.columns and '7个月涨跌幅' in df.columns:
            df[label_col] = (df['7个月涨跌幅'] > threshold / 100).astype(int)
    if label_col not in df.columns:
        print(f'❌ 找不到标签列 {label_col}')
        return None, None

    y = df[label_col]
    valid = y.notna()
    df = df[valid].reset_index(drop=True)
    y = y[valid].reset_index(drop=True)

    available = []
    for c in FEATURE_COLUMNS + MYSQL_RATIO_COLS:
        if c in df.columns:
            available.append(c)
    cat_available = [c for c in CATEGORICAL_COLS if c in df.columns]

    X = df[available + cat_available].copy()
    for c in available:
        X[c] = pd.to_numeric(X[c], errors='coerce')
        X[c] = X[c].fillna(X[c].median())
    for c in cat_available:
        X[c] = X[c].fillna('未知').astype(str)

    print(f'特征数: {len(available) + len(cat_available)} (手工选择)')
    print(f'盈利占比: {y.mean()*100:.1f}%')
    print(f'样本数: {len(y)} (盈利{y.sum():.0f} + 亏损{(1-y).sum():.0f})')

    if y.sum() < 5 or (1 - y).sum() < 5:
        print(f'⚠️ 样本太少或单一类别')
        return None, None

    return X, y


def train_lightgbm(X, y, output_dir, suffix=''):
    """LightGBM分类模型"""
    from sklearn.model_selection import StratifiedKFold, cross_val_score, cross_val_predict
    from sklearn.metrics import classification_report
    import lightgbm as lgb

    print('\n' + '='*60)
    print(f'LightGBM 模型（精准度高）{"["+suffix+"]" if suffix else ""}')
    print('='*60)

    model = lgb.LGBMClassifier(**LGB_PARAMS)

    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    auc = cross_val_score(model, X, y, cv=cv, scoring='roc_auc')
    print(f'5折AUC: {auc.mean():.3f} ± {auc.std():.3f}')

    y_pred = cross_val_predict(model, X, y, cv=cv)
    print(f'\n分类报告:')
    print(classification_report(y, y_pred, target_names=['亏损', '盈利']))

    # 全量拟合 + 早停防过拟合(切20%做 held-out val; 失败则退回全量无早停)
    from sklearn.model_selection import train_test_split as _tts
    try:
        _a, _b, _ya, _yb = _tts(X, y, test_size=0.2, random_state=42, stratify=y)
        model.fit(_a, _ya, eval_set=[(_b, _yb)], callbacks=[lgb.early_stopping(20, verbose=False)])
    except Exception:
        model.fit(X, y)
    model_file = os.path.join(output_dir, f'lgb_classifier{suffix}.txt')
    model.booster_.save_model(model_file)

    # 特征重要性
    importance = pd.DataFrame({
        'feature': X.columns, 'importance': model.feature_importances_
    }).sort_values('importance', ascending=False)
    print(f'\n特征重要性 Top 10:')
    print(importance.head(10).to_string(index=False))
    importance.to_csv(os.path.join(output_dir, f'lgb_feature_importance{suffix}.csv'), index=False)

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
        plt.title(f'LightGBM 特征重要性{(" ("+suffix+")") if suffix else ""}', fontsize=14)
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, f'lgb_feature_importance{suffix}.png'), dpi=150)
        plt.close()
    except Exception:
        pass

    return model, auc.mean()


def train_logistic_regression(X, y, output_dir, suffix=''):
    """逻辑回归模型（可读性高，评分卡风格）"""
    from sklearn.model_selection import StratifiedKFold, cross_val_score, cross_val_predict
    from sklearn.metrics import classification_report
    from sklearn.linear_model import LogisticRegression
    from sklearn.preprocessing import StandardScaler

    print('\n' + '='*60)
    print(f'逻辑回归模型（可读性高，评分卡风格）{"["+suffix+"]" if suffix else ""}')
    print('='*60)

    # 全量模式已经是纯数值，无需额外处理
    X_lr = X.copy()
    X_lr = X_lr.replace([np.inf, -np.inf], 0).fillna(0)

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

    # 保存模型+scaler+特征列表+median
    medians = {c: float(X_lr[c].median()) for c in X_lr.columns}
    pkl_file = os.path.join(output_dir, f'lr_classifier{suffix}.pkl')
    with open(pkl_file, 'wb') as f:
        pickle.dump({
            'model': model,
            'scaler': scaler,
            'features': list(X_lr.columns),
            'medians': medians,
        }, f)

    # 系数表（评分卡）
    coef_df = pd.DataFrame({
        '特征': X_lr.columns,
        '标准化系数': model.coef_[0],
    }).sort_values('标准化系数', ascending=False)
    coef_df['影响'] = coef_df['标准化系数'].apply(
        lambda x: '↑正相关（高分→盈利）' if x > 0 else '↓负相关（高分→亏损）'
    )
    coef_df['影响强度'] = coef_df['标准化系数'].abs().apply(
        lambda x: '★★★' if x > 0.3 else ('★★' if x > 0.15 else '★')
    )
    coef_df = coef_df.rename(columns={'标准化系数': '系数(标准化后)'})

    print(f'\n逻辑回归系数表（评分卡）:')
    print(coef_df.head(15).to_string(index=False))
    coef_df.to_csv(os.path.join(output_dir, f'lr_coefficients{suffix}.csv'), index=False)

    return model, scaler, auc.mean()


def evaluate_oot(X, y, year, split_year=2024):
    """时序 out-of-time: train(报价日年<=split_year) / test(>split_year)。真实泛化指标。

    随机 5 折 CV(shuffle) 会把未来样本混进训练折 → 虚高; OOT 才是主指标。
    LGB 用 LGB_PARAMS(降复杂度)+train内早停; LR 同 train_logistic_regression。
    返回 dict(lgb_oot_auc, lr_oot_auc, train_n, test_n, train_pos, test_pos) 或 None。
    """
    from sklearn.metrics import roc_auc_score
    from sklearn.model_selection import train_test_split
    from sklearn.preprocessing import StandardScaler
    from sklearn.linear_model import LogisticRegression
    import lightgbm as lgb
    tr = (year <= split_year).values
    te = (year > split_year).values
    if te.sum() < 10 or tr.sum() < 50:
        print('  ⚠ OOT: 时序切分样本不足, 跳过')
        return None
    Xtr = X[tr].replace([np.inf, -np.inf], 0).fillna(0)
    Xte = X[te].replace([np.inf, -np.inf], 0).fillna(0)
    ytr, yte = y[tr], y[te]
    # LGB(降复杂度 + train内早停)
    m = lgb.LGBMClassifier(**LGB_PARAMS)
    a, b, ya, yb = train_test_split(Xtr, ytr, test_size=0.2, random_state=42, stratify=ytr)
    try:
        m.fit(a, ya, eval_set=[(b, yb)], callbacks=[lgb.early_stopping(20, verbose=False)])
    except Exception:
        m.fit(Xtr, ytr)
    lgb_oot = roc_auc_score(yte, m.predict_proba(Xte)[:, 1])
    # LR
    sc = StandardScaler().fit(Xtr)
    lr = LogisticRegression(C=1.0, penalty='l2', max_iter=1000, random_state=42)
    lr.fit(sc.transform(Xtr), ytr)
    lr_oot = roc_auc_score(yte, lr.predict_proba(sc.transform(Xte))[:, 1])
    res = dict(lgb_oot_auc=float(lgb_oot), lr_oot_auc=float(lr_oot),
               train_n=int(tr.sum()), test_n=int(te.sum()),
               train_pos=float(ytr.mean()), test_pos=float(yte.mean()))
    print(f'\n{"="*60}\n[主指标] 时序 OOT (train≤{split_year} / test>{split_year})')
    print(f'  train n={res["train_n"]}(正{res["train_pos"]*100:.1f}%) | test n={res["test_n"]}(正{res["test_pos"]*100:.1f}%)')
    print(f'  LightGBM OOT AUC = {res["lgb_oot_auc"]:.3f}')
    print(f'  逻辑回归  OOT AUC = {res["lr_oot_auc"]:.3f}')
    print(f'  (随机5折CV因shuffle混入未来样本会虚高, 以此 OOT 为准)\n{"="*60}')
    return res


def main():
    parser = argparse.ArgumentParser(description='定增盈利预测 - 双模型训练')
    parser.add_argument('features_path', help='features.parquet 或 features_derived.parquet 路径')
    parser.add_argument('--threshold', type=float, default=-10, help='盈利阈值(%%)，默认-10')
    parser.add_argument('--label', default=None,
                        help='显式标签列(如 标签_盈利_-10_3m / 标签_极性_灰度剔除_7m); '
                             '默认按 --threshold 合成 标签_盈利_{threshold}(=7m)')
    parser.add_argument('--split-year', type=int, default=2024, help='OOT 时序切分年(train<=Y / test>Y)')
    parser.add_argument('--dataset-version', default=None, help='DB 快照版本(记入 registry, 可追溯训练数据)')
    parser.add_argument('--no-set-current', action='store_true', help='注册但不设为 current(扫描用)')
    parser.add_argument('--classic', action='store_true',
                        help='使用手工选择特征（64个），默认使用全部数值特征')
    args = parser.parse_args()

    output_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'output')
    os.makedirs(output_dir, exist_ok=True)

    # 加载
    df = load_features(args.features_path)

    label_col = args.label  # None → 按 threshold 合成 标签_盈利_{threshold}
    if args.classic:
        X, y = prepare_features_classic(df, args.threshold, label_col=label_col)
        suffix = ''
    else:
        X, y = prepare_features_full(df, args.threshold, label_col=label_col)
        suffix = '_full'

    if X is None:
        sys.exit(1)

    # 训练LightGBM
    lgb_model, lgb_auc = train_lightgbm(X, y, output_dir, suffix=suffix)

    # 训练逻辑回归
    lr_model, lr_scaler, lr_auc = train_logistic_regression(X, y, output_dir, suffix=suffix)

    # 时序 OOT(真实泛化, 主指标)
    _label = label_col or f'标签_盈利_{int(args.threshold)}'
    _valid = df[_label].notna() if _label in df.columns else pd.Series([True] * len(df))
    year = pd.to_numeric(pd.to_numeric(df.loc[_valid, '报价日'], errors='coerce')
                         .astype('Int64').astype(str).str[:4], errors='coerce').reset_index(drop=True)
    oot = evaluate_oot(X, y, year, split_year=args.split_year)

    # 保存全量模型的 meta（特征列表 + median，供 predict_profitability 使用）
    if not args.classic:
        meta = {
            'features': list(X.columns),
            'medians': {c: float(X[c].median()) for c in X.columns},
            'threshold': args.threshold,
        }
        meta_path = os.path.join(output_dir, 'lgb_full_meta.json')
        with open(meta_path, 'w') as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print(f'\n模型meta已保存: {meta_path}')

    # 模型对比
    print('\n' + '='*60)
    print('模型对比')
    print('='*60)
    print(f'{"模型":<20} {"特征数":<8} {"AUC":<10} {"优势"}')
    print(f'{"LightGBM":<20} {X.shape[1]:<8} {lgb_auc:<10.3f} 精准度高，捕捉非线性')
    print(f'{"逻辑回归":<20} {X.shape[1]:<8} {lr_auc:<10.3f} 可读性强，系数=评分卡')

    # 保存评估报告
    report_path = os.path.join(output_dir, 'evaluation_report.txt')
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(f'定增盈利预测模型评估报告\n')
        f.write(f'盈利阈值: {args.threshold}%\n')
        f.write(f'特征模式: {"手工选择(64特征)" if args.classic else "全量数值特征"}\n')
        f.write(f'特征数: {X.shape[1]}\n')
        f.write(f'样本数: {len(y)}, 盈利占比: {y.mean()*100:.1f}%\n\n')
        f.write(f'LightGBM AUC(5折CV): {lgb_auc:.3f}\n')
        f.write(f'逻辑回归 AUC(5折CV): {lr_auc:.3f}\n')
        if oot:
            f.write(f'\n[主指标] 时序 OOT (train≤2024 / test≥2025):\n')
            f.write(f'  LightGBM OOT AUC: {oot["lgb_oot_auc"]:.3f}\n')
            f.write(f'  逻辑回归  OOT AUC: {oot["lr_oot_auc"]:.3f}\n')
            f.write(f'  train n={oot["train_n"]}(正{oot["train_pos"]*100:.1f}%) / test n={oot["test_n"]}(正{oot["test_pos"]*100:.1f}%)\n')
        f.write(f'\n逻辑回归系数表见: lr_coefficients{suffix}.csv\n')
        f.write(f'LightGBM特征重要性见: lgb_feature_importance{suffix}.csv\n')

    # ═══════════════════════════════════════════════
    # 版本归档: 复制到版本子目录
    # ═══════════════════════════════════════════════
    import shutil
    date_str = datetime.now().strftime('%Y%m%d_%H%M')
    feat_tag = f'{X.shape[1]}feat'
    ks_tag = f'auc{lgb_auc:.2f}'.replace('.', '')
    version_dir = os.path.join(output_dir, f'v_{date_str}_{feat_tag}_{ks_tag}')
    os.makedirs(version_dir, exist_ok=True)

    # 需要归档的文件
    archive_files = [
        f'lgb_classifier{suffix}.txt',
        f'lgb_feature_importance{suffix}.csv',
        f'lgb_feature_importance{suffix}.png',
        f'lr_classifier{suffix}.pkl',
        f'lr_coefficients{suffix}.csv',
        'evaluation_report.txt',
    ]
    if not args.classic:
        archive_files.append('lgb_full_meta.json')

    for fname in archive_files:
        src = os.path.join(output_dir, fname)
        if os.path.exists(src):
            shutil.copy2(src, os.path.join(version_dir, fname))

    # 写入版本说明
    version_md = os.path.join(version_dir, 'VERSION.md')
    with open(version_md, 'w', encoding='utf-8') as f:
        f.write(f'# 模型版本: v_{date_str}_{feat_tag}_{ks_tag}\n\n')
        f.write(f'**训练时间**: {datetime.now().strftime("%Y-%m-%d %H:%M")}\n')
        f.write(f'**特征数**: {X.shape[1]}\n')
        f.write(f'**特征模式**: {"手工选择(64特征)" if args.classic else "全量数值特征"}\n')
        f.write(f'**训练样本**: {len(y)}, 盈利占比: {y.mean()*100:.1f}%\n')
        f.write(f'**盈利阈值**: {args.threshold}%\n\n')
        f.write(f'## 指标\n')
        f.write(f'| 模型 | 5折CV AUC | OOT AUC |\n|------|---------|---------|\n')
        f.write(f'| LightGBM | {lgb_auc:.3f} | {oot["lgb_oot_auc"]:.3f} |\n' if oot else f'| LightGBM | {lgb_auc:.3f} | - |\n')
        f.write(f'| 逻辑回归 | {lr_auc:.3f} | {oot["lr_oot_auc"]:.3f} |\n\n' if oot else f'| 逻辑回归 | {lr_auc:.3f} | - |\n\n')
        f.write(f'## LightGBM 参数(降复杂度防过拟合 + 早停)\n')
        f.write(f'- n_estimators={LGB_PARAMS["n_estimators"]}, max_depth={LGB_PARAMS["max_depth"]}, '
                f'num_leaves={LGB_PARAMS["num_leaves"]}, min_child_samples={LGB_PARAMS["min_child_samples"]}, '
                f'lr={LGB_PARAMS["learning_rate"]}, +early_stopping\n\n')
        f.write(f'## 文件清单\n')
        for fname in archive_files:
            if os.path.exists(os.path.join(version_dir, fname)):
                f.write(f'- {fname}\n')

    # ═══════════════════════════════════════════════
    # 注册到 model_registry（predict_profitability 从 registry 读 current 版本）
    # ═══════════════════════════════════════════════
    if not args.classic:
        try:
            from deploy.model_registry import register_version
            version_name = os.path.basename(version_dir)
            archived = [f for f in archive_files
                        if os.path.exists(os.path.join(version_dir, f))]
            register_version(
                model_type='full',
                version=version_name,
                dir=version_dir,
                metrics={'lgb_auc': float(lgb_auc), 'lr_auc': float(lr_auc),
                         **({'lgb_oot_auc': oot['lgb_oot_auc'], 'lr_oot_auc': oot['lr_oot_auc']} if oot else {})},
                n_features=X.shape[1],
                threshold=args.threshold,
                n_samples=len(y),
                positive_rate=float(y.mean()),
                files=archived,
                note='全量数值特征 LGB+LR',
                set_current=not args.no_set_current,
                label_config=label_col or f'7m_{int(args.threshold)}',
                dataset_version=args.dataset_version,
            )
            cur_tag = '' if args.no_set_current else ' 已注册为 full 当前版本'
            print(f'   registry: 已注册 → {version_name}{cur_tag}')
        except Exception as e:
            print(f'   ⚠ registry 注册失败(不影响模型文件): {e}')

    print(f'\n✅ 完成! 输出目录: {output_dir}')
    print(f'   版本归档: {version_dir}')


if __name__ == '__main__':
    main()
