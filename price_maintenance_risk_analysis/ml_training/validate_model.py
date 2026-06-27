#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增模型验证脚本 - 对询转项目做完整特征提取+预测+校验

流程:
  1. 从Excel提取测试样本基础信息
  2. 完整运行 export_features (DB特征)
  3. 完整运行 derive_features (衍生特征)
  4. 用训练集的median填充缺失特征
  5. 加载模型预测，对比实际解禁浮盈

用法:
    python ml_training/validate_model.py <test_excel_path>
"""

import sys
import os
import argparse
import json
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)


def main():
    parser = argparse.ArgumentParser(description='定增模型验证')
    parser.add_argument('test_excel', help='询转项目Excel路径')
    args = parser.parse_args()

    print('=' * 70)
    print('定增模型验证 - 询转项目测试集')
    print('=' * 70)

    # ====== 1. 加载测试数据 ======
    print('\n1. 加载测试数据...')
    raw = pd.read_excel(args.test_excel, sheet_name='已完成发行询转项目')
    print(f'  原始 {len(raw)} 条')

    # 构造 scored 风格的 DataFrame（供 export_features 加载）
    scored = pd.DataFrame()
    scored['股票代码'] = raw['股票代码'].astype(str)
    scored['股票简称'] = raw['项目名称'].astype(str)
    # 报价日
    dt = pd.to_datetime(raw['报价日期'], errors='coerce')
    scored['报价日_excel'] = dt.dt.strftime('%Y%m%d')
    scored['报价日'] = dt.dt.strftime('%Y%m%d').apply(lambda x: x if pd.notna(x) and len(x) == 8 else None)
    scored['报价日价格'] = pd.to_numeric(raw['报价前日市价'], errors='coerce')

    # 行业
    if '申万行业分类' in raw.columns:
        parts = raw['申万行业分类'].astype(str).str.split('--', expand=True)
        scored['一级行业'] = parts[0] if parts is not None and 0 in parts.columns else '未知'
        scored['二级行业'] = parts[1] if parts is not None and 1 in parts.columns else '未知'
        scored['三级行业'] = parts[2] if parts is not None and 2 in parts.columns else '未知'

    # 标签: 解禁浮盈
    target = pd.to_numeric(raw['解禁浮盈'], errors='coerce')
    scored['7个月涨跌幅'] = target * 100  # 转百分比，和训练集一致
    scored['标签_盈利_0'] = (target > 0).astype(int)
    scored['标签_盈利_-10'] = (target > -0.10).astype(int)
    scored['标签_盈利_-20'] = (target > -0.20).astype(int)

    # 子场景全部置空（测试集没有筛选结果）
    for c in ['市场指数', '行业PE', '个股PE', 'DCF估值', '修正PE估值',
              '参数构造', '蒙特卡洛', '反向推算']:
        scored[c] = np.nan
    scored['定增决策'] = np.nan
    scored['有效阈值数'] = np.nan

    # 评分全部置空
    for prefix in ['总分', '评级', '盈利能力', '成长能力']:
        for suffix in ['T-4', 'T-3', 'T-2', 'T-1', 'T', '斜率', '趋势']:
            scored[f'{prefix}_{suffix}'] = np.nan
    scored['综合趋势'] = np.nan

    # 去掉没有标签的行
    valid = scored['标签_盈利_-10'].notna()
    scored = scored[valid].reset_index(drop=True)
    print(f'  有效样本: {len(scored)} 条, 盈利(>0): {(scored["标签_盈利_0"]==1).sum()} ({(scored["标签_盈利_0"]==1).mean()*100:.1f}%)')

    # 保存原始信息
    meta = pd.DataFrame({
        '股票代码': scored['股票代码'],
        '项目名称': raw.loc[valid, '项目名称'].values,
        '中标价': pd.to_numeric(raw.loc[valid, '中标价'], errors='coerce').values,
        '折扣': pd.to_numeric(raw.loc[valid, '折扣'], errors='coerce').values,
        '解禁浮盈': target[valid].values,
        '解禁日': pd.to_datetime(raw.loc[valid, '解禁日（预计）'], errors='coerce').values,
    })

    # ====== 2. 完整 DB 特征提取 (复用 export_features) ======
    print('\n2. 完整DB特征提取...')
    from features.export_features import load_db_features, load_financial_ratios
    sample_keys = list(zip(scored['股票代码'].tolist(), scored['报价日_excel'].tolist()))
    db_feats = load_db_features(sample_keys)
    db_feat_cols = [c for c in db_feats.columns if c != '股票代码']
    # 行对齐合并（先去掉scored中已有的db列，避免重复）
    dup_cols = [c for c in db_feat_cols if c in scored.columns]
    if dup_cols:
        db_feats = db_feats.drop(columns=dup_cols)
        db_feat_cols = [c for c in db_feats.columns if c != '股票代码']
    df = pd.concat([scored.reset_index(drop=True), db_feats[db_feat_cols].reset_index(drop=True)], axis=1)

    # 财务比率(financial_indicators) PIT —— 与 export_features.main() 同源(按报价日回溯)
    try:
        ratio_feats = load_financial_ratios(sample_keys)
        if ratio_feats is not None and not ratio_feats.empty:
            rcols = [c for c in ratio_feats.columns if c not in df.columns]
            if rcols:
                df = pd.concat([df.reset_index(drop=True), ratio_feats[rcols].reset_index(drop=True)], axis=1)
                print(f'  财务比率 PIT(financial_indicators): +{len(rcols)} 列')
    except Exception as e:
        print(f'  财务比率跳过: {e}')

    matched_price = df.get('当前价', pd.Series()).notna().sum()
    print(f'  行情匹配: {matched_price}/{len(df)}')

    # 清理类型
    str_keep = {'股票代码', '股票简称', '最终结论', '一级行业', '二级行业', '三级行业',
                '定价方式', '定增决策', '行业代码', '行业名称'}
    for c in df.columns:
        if c in str_keep:
            df[c] = df[c].astype(str)
        elif df[c].dtype == object or str(df[c].dtype) == 'category':
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # ====== 3. 完整衍生特征 (复用 derive_features) ======
    print('\n3. 完整衍生特征...')
    from features.derive_features import (
        derive_fcf_growth_rates, derive_fcf_cross_metrics,
        derive_financial_score_deltas, derive_valuation_relative,
        derive_market_momentum, derive_industry_valuation_growth,
        derive_market_index_features
    )
    for func in [derive_fcf_growth_rates, derive_fcf_cross_metrics,
                 derive_financial_score_deltas, derive_valuation_relative,
                 derive_market_momentum]:
        df = func(df)
    try:
        df = derive_industry_valuation_growth(df)
        df = derive_market_index_features(df)
    except Exception as e:
        print(f'  DB衍生跳过: {e}')

    df = df.replace([np.inf, -np.inf], np.nan)
    print(f'  最终特征: {df.shape[1]} 列')

    # ====== 4. 在训练集上训练纯数值模型 ======
    print('\n4. 训练纯数值特征模型(用于验证)...')
    import lightgbm as lgb
    from sklearn.model_selection import StratifiedKFold, cross_val_score

    train_path = os.path.join(SCRIPT_DIR, 'data', 'features_derived.parquet')
    train_df = pd.read_parquet(train_path)
    train_y = train_df['标签_盈利_-10'].dropna()
    train_df = train_df.loc[train_y.index].reset_index(drop=True)
    train_y = train_y.reset_index(drop=True)

    # 只取数值列，排除标签和目标
    exclude = {'7个月涨跌幅', '7个月后价格', '标签_盈利_0', '标签_盈利_-10', '标签_盈利_-20'}
    num_cols = [c for c in train_df.select_dtypes(include=[np.number]).columns if c not in exclude]

    train_X = train_df[num_cols].copy()
    for c in num_cols:
        train_X[c] = pd.to_numeric(train_X[c], errors='coerce').fillna(train_X[c].median())
    train_X = train_X.replace([np.inf, -np.inf], 0)

    # 用和之前最优的参数训练
    val_model = lgb.LGBMClassifier(
        n_estimators=300, max_depth=5, learning_rate=0.03, num_leaves=31,
        is_unbalance=True, subsample=0.8, colsample_bytree=0.8,
        reg_alpha=0.1, reg_lambda=0.1, random_state=42, verbose=-1
    )
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    cv_auc = cross_val_score(val_model, train_X, train_y, cv=cv, scoring='roc_auc')
    print(f'  训练集CV AUC: {cv_auc.mean():.3f} ± {cv_auc.std():.3f}')

    # 全量训练
    val_model.fit(train_X, train_y)
    model_features = list(train_X.columns)
    print(f'  模型特征: {len(model_features)} 个(纯数值)')

    # ====== 5. 构建测试特征矩阵并预测 ======
    print('\n5. 构建测试特征矩阵并预测...')
    X_test = pd.DataFrame(index=df.index)
    missing_list = []
    for feat in model_features:
        if feat in df.columns:
            X_test[feat] = pd.to_numeric(df[feat], errors='coerce')
        else:
            X_test[feat] = np.nan
            missing_list.append(feat)

    # 用训练集 median 填充缺失
    train_medians = train_X.median()
    for feat in model_features:
        X_test[feat] = X_test[feat].fillna(train_medians.get(feat, 0))
    X_test = X_test.replace([np.inf, -np.inf], 0)

    if missing_list:
        print(f'  ⚠️ 缺少 {len(missing_list)} 个特征(用训练集median填充): {missing_list[:8]}...')

    # 特征覆盖率统计
    total_cells = len(X_test) * len(model_features)
    non_null_before_fill = sum(df[f].notna().sum() for f in model_features if f in df.columns)
    print(f'  特征覆盖率(填充前): {non_null_before_fill/total_cells*100:.1f}%')

    # 预测
    df['预测盈利概率'] = val_model.predict_proba(X_test)[:, 1]
    print(f'  预测完成, 概率范围: [{df["预测盈利概率"].min():.3f}, {df["预测盈利概率"].max():.3f}]')

    # ====== 6. 评估 ======
    from sklearn.metrics import roc_auc_score, confusion_matrix

    def calc_ks(y_true, y_proba):
        pos = y_proba[y_true == 1]
        neg = y_proba[y_true == 0]
        thresholds = np.sort(np.unique(y_proba))
        if len(thresholds) < 2:
            return 0.0
        return max(abs((pos >= t).mean() - (neg >= t).mean()) for t in thresholds)

    for threshold_label, col in [('盈利>0', '标签_盈利_0'), ('盈利>-10%', '标签_盈利_-10')]:
        y_true = df[col]
        y_proba = df['预测盈利概率']
        valid_mask = y_true.notna() & y_proba.notna() & ~np.isinf(y_proba)
        if valid_mask.sum() < 20:
            continue

        yt = y_true[valid_mask].astype(int)
        yp = y_proba[valid_mask]
        auc = roc_auc_score(yt, yp)
        ks = calc_ks(yt, yp)

        print(f'\n{"=" * 70}')
        print(f'验证结果: {threshold_label}')
        print(f'{"=" * 70}')
        print(f'  AUC = {auc:.3f}   KS = {ks:.3f}')
        print(f'  样本: {len(yt)}, 盈利: {yt.sum()} ({yt.mean()*100:.1f}%)')

        # 混淆矩阵 (0.5阈值)
        y_pred = (yp > 0.5).astype(int)
        cm = confusion_matrix(yt, y_pred)
        print(f'\n  混淆矩阵(阈值=0.5):')
        print(f'  {"":>10s} {"预测亏损":>8s} {"预测盈利":>8s}')
        print(f'  {"实际亏损":>10s} {cm[0,0]:8d} {cm[0,1]:8d}')
        print(f'  {"实际盈利":>10s} {cm[1,0]:8d} {cm[1,1]:8d}')
        precision = cm[1,1] / max(cm[1,1]+cm[0,1], 1)
        recall = cm[1,1] / max(cm[1,1]+cm[1,0], 1)
        print(f'  精确率(Precision)={precision:.3f}  召回率(Recall)={recall:.3f}')

        # 分位分组
        df_eval = pd.DataFrame({'y': yt.values, 'p': yp.values})
        try:
            df_eval['decile'] = pd.qcut(df_eval['p'], 10, duplicates='drop')
            print(f'\n  概率分位 → 实际盈利率:')
            print(f'  {"分位":>35s} {"数量":>5s} {"盈利":>5s} {"盈利率":>7s} {"平均概率":>8s}')
            for idx, grp in df_eval.groupby('decile', observed=True):
                n = len(grp)
                pos_n = grp['y'].sum()
                print(f'  {str(idx):>35s} {n:5d} {pos_n:5.0f} {pos_n/n*100:6.1f}% {grp["p"].mean():8.3f}')
        except Exception:
            pass

    # ====== 7. 解禁浮盈 vs 预测概率 ======
    print(f'\n{"=" * 70}')
    print(f'解禁浮盈(连续值) vs 预测概率')
    print(f'{"=" * 70}')
    ret = meta['解禁浮盈'].astype(float)
    prob = df['预测盈利概率']
    valid = ret.notna() & prob.notna()
    ret_v = ret[valid].reset_index(drop=True)
    prob_v = prob[valid].reset_index(drop=True)

    # 相关性
    corr = ret_v.corr(prob_v)
    print(f'  相关系数(概率 vs 浮盈): {corr:.3f}')

    print(f'\n  {"概率分组":>20s} {"样本":>5s} {"平均浮盈":>10s} {"盈利占比":>8s} {"中标折扣":>8s}')
    discount = meta.loc[valid, '折扣'].reset_index(drop=True)
    for lo in np.arange(0, 1.0, 0.1):
        hi = lo + 0.1
        mask = (prob_v >= lo) & (prob_v < hi)
        n = mask.sum()
        if n < 3:
            continue
        avg_ret = ret_v[mask].mean()
        profit_rate = (ret_v[mask] > 0).mean() * 100
        avg_disc = discount[mask].mean()
        print(f'  [{lo:.1f}, {hi:.1f})       {n:5d} {avg_ret:+9.2%} {profit_rate:7.1f}% {avg_disc:8.3f}')

    # ====== 8. 保存 ======
    output_path = os.path.join(SCRIPT_DIR, 'output', 'validation_result.csv')
    save_cols = ['股票代码', '项目名称', '中标价', '折扣', '解禁浮盈', '解禁日',
                 '预测盈利概率']
    save_df = meta.copy()
    save_df['预测盈利概率'] = df['预测盈利概率'].values
    save_df['标签_盈利_0'] = (save_df['解禁浮盈'] > 0).astype(int)
    save_df['标签_盈利_-10'] = (save_df['解禁浮盈'] > -0.10).astype(int)
    save_df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f'\n✅ 验证结果已保存: {output_path}')


if __name__ == '__main__':
    main()
