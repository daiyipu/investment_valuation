#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增盈利概率预测 - 加载训练好的双模型对新样本预测

模型: 全量数值特征版本（320特征）
  - lgb_classifier_full.txt  (LightGBM, CV AUC=0.746)
  - lr_classifier_full.pkl   (逻辑回归, CV AUC=0.655)

支持两种用法:
  1. 独立运行: python ml_training/predict_profitability.py <scored_excel> [--output result.xlsx]
  2. 被调用:   from predict_profitability import predict; df = predict(scored_excel_path)

输出两个模型的盈利概率:
  - 盈利概率_LightGBM
  - 盈利概率_逻辑回归
"""

import sys
import os
import json
import warnings
import pickle
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)


def predict(scored_excel_path):
    """
    对 scored Excel 中的每只股票预测盈利概率

    流程:
      1. 从 Excel 提取财务评分/子场景等特征
      2. 从 MySQL 加载行情/估值/FCF 等DB特征
      3. 计算衍生特征（FCF增长率、交叉比率、评分变动、估值相对、行情动量、行业/大盘）
      4. 加载全量特征模型的 meta（特征列表 + median）
      5. LightGBM + 逻辑回归 双模型预测

    Args:
        scored_excel_path: 已评分的 Excel 文件路径（batch_screen_and_score 输出）

    Returns:
        DataFrame: columns = [股票代码, 盈利概率_LightGBM, 盈利概率_逻辑回归]
    """
    import lightgbm as lgb
    from export_features import load_db_features, load_scored_features, load_financial_ratios, load_financial_statements
    from derive_features import (
        derive_fcf_growth_rates, derive_fcf_cross_metrics,
        derive_financial_score_deltas, derive_valuation_relative,
        derive_market_momentum, derive_industry_valuation_growth,
        derive_market_index_features,
    )

    output_dir = os.path.join(SCRIPT_DIR, 'output')

    # ═══════════════════════════════════════════════
    # 1. 从 Excel 提取特征
    # ═══════════════════════════════════════════════
    print('\n  [ML-1] 加载 Excel 特征...')
    scored = load_scored_features(scored_excel_path)
    stock_codes = scored['股票代码'].tolist()
    print(f'    样本: {len(scored)} 条')

    # ═══════════════════════════════════════════════
    # 2. 从 DB 加载特征
    # ═══════════════════════════════════════════════
    print('  [ML-2] 加载 DB 特征...')
    sample_keys = []
    for _, row in scored.iterrows():
        code = row['股票代码']
        issue_date = str(row.get('报价日_excel', '')).replace('.0', '').strip()
        if not issue_date or issue_date == 'nan' or len(issue_date) < 8:
            issue_date = None
        sample_keys.append((code, issue_date))

    db_feats = load_db_features(sample_keys)
    matched_price = db_feats.get('当前价', pd.Series()).notna().sum()
    print(f'    行情匹配: {matched_price}/{len(scored)}')

    # 财务比率 + 三表（可选，失败不中断）
    try:
        ratio_feats = load_financial_ratios(stock_codes)
        stmt_feats = load_financial_statements(stock_codes)
    except Exception as e:
        print(f'    财务比率跳过: {e}')
        ratio_feats = pd.DataFrame()
        stmt_feats = pd.DataFrame()

    # ═══════════════════════════════════════════════
    # 3. 合并
    # ═══════════════════════════════════════════════
    scored = scored.reset_index(drop=True)
    db_feat_cols = [c for c in db_feats.columns if c != '股票代码']
    df = pd.concat([scored, db_feats[db_feat_cols].reset_index(drop=True)], axis=1)

    if not ratio_feats.empty:
        df = df.merge(ratio_feats, on='股票代码', how='left')
    if not stmt_feats.empty:
        df = df.merge(stmt_feats, on='股票代码', how='left')

    # 清理类型
    str_keep = {'股票代码', '股票简称', '最终结论', '一级行业', '二级行业', '三级行业',
                '定价方式', '定增决策', '行业代码', '行业名称'}
    for c in df.columns:
        if c in str_keep:
            df[c] = df[c].astype(str)
        elif df[c].dtype == object or str(df[c].dtype) == 'category':
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # ═══════════════════════════════════════════════
    # 4. 衍生特征
    # ═══════════════════════════════════════════════
    print('  [ML-3] 衍生特征...')
    for func in [derive_fcf_growth_rates, derive_fcf_cross_metrics,
                 derive_financial_score_deltas, derive_valuation_relative,
                 derive_market_momentum]:
        df = func(df)
    try:
        df = derive_industry_valuation_growth(df)
        df = derive_market_index_features(df)
    except Exception as e:
        print(f'    行业/大盘衍生跳过: {e}')

    df = df.replace([np.inf, -np.inf], np.nan)
    print(f'    最终特征: {df.shape[1]} 列')

    # ═══════════════════════════════════════════════
    # 5. 加载全量特征模型的 meta
    # ═══════════════════════════════════════════════
    print('  [ML-4] 加载模型 meta...')
    with open(os.path.join(output_dir, 'lgb_full_meta.json'), 'r') as f:
        meta = json.load(f)
    model_features = meta['features']
    train_medians = meta['medians']
    print(f'    模型特征: {len(model_features)} 个')

    # ═══════════════════════════════════════════════
    # 6. 构建预测特征矩阵
    # ═══════════════════════════════════════════════
    print('  [ML-5] 构建预测特征矩阵...')
    X_pred = pd.DataFrame(index=df.index)
    missing_list = []
    for feat in model_features:
        if feat in df.columns:
            X_pred[feat] = pd.to_numeric(df[feat], errors='coerce')
        else:
            X_pred[feat] = np.nan
            missing_list.append(feat)

    # 用训练集 median 填充
    for feat in model_features:
        X_pred[feat] = X_pred[feat].fillna(train_medians.get(feat, 0))
    X_pred = X_pred.replace([np.inf, -np.inf], 0).fillna(0)

    if missing_list:
        print(f'    ⚠️ 缺少 {len(missing_list)} 个特征(median填充): {missing_list[:5]}...')

    # 特征覆盖率
    total_cells = len(X_pred) * len(model_features)
    non_null = sum(df[f].notna().sum() for f in model_features if f in df.columns)
    print(f'    特征覆盖率(填充前): {non_null/total_cells*100:.1f}%')

    # ═══════════════════════════════════════════════
    # 7. LightGBM 预测
    # ═══════════════════════════════════════════════
    print('  [ML-6] LightGBM 预测...')
    booster = lgb.Booster(model_file=os.path.join(output_dir, 'lgb_classifier_full.txt'))
    lgb_feature_names = booster.feature_name()

    # 对齐特征
    X_lgb = pd.DataFrame(index=X_pred.index)
    for feat in lgb_feature_names:
        X_lgb[feat] = X_pred[feat] if feat in X_pred.columns else 0
    X_lgb = X_lgb.astype(float).values

    lgb_proba = booster.predict(X_lgb)
    print(f'    概率范围: [{lgb_proba.min():.3f}, {lgb_proba.max():.3f}]')

    # ═══════════════════════════════════════════════
    # 8. 逻辑回归 预测
    # ═══════════════════════════════════════════════
    print('  [ML-7] 逻辑回归 预测...')
    with open(os.path.join(output_dir, 'lr_classifier_full.pkl'), 'rb') as f:
        lr_data = pickle.load(f)
    lr_model = lr_data['model']
    lr_scaler = lr_data['scaler']
    lr_feature_names = lr_data['features']

    # 对齐特征
    X_lr = pd.DataFrame(index=X_pred.index)
    for feat in lr_feature_names:
        X_lr[feat] = X_pred[feat] if feat in X_pred.columns else 0
    X_lr = X_lr.replace([np.inf, -np.inf], 0).fillna(0)
    X_lr_scaled = lr_scaler.transform(X_lr.values)
    lr_proba = lr_model.predict_proba(X_lr_scaled)[:, 1]
    print(f'    概率范围: [{lr_proba.min():.3f}, {lr_proba.max():.3f}]')

    # ═══════════════════════════════════════════════
    # 9. 返回
    # ═══════════════════════════════════════════════
    result = pd.DataFrame({
        '股票代码': df['股票代码'],
        '盈利概率_LightGBM': lgb_proba,
        '盈利概率_逻辑回归': lr_proba,
    })

    print(f'  ✅ ML预测完成: {len(result)} 条\n')
    return result


def write_to_excel(scored_excel_path, result_df, output_path=None):
    """将预测结果写入 Excel（追加到最后两列，按行序一一对应写入）

    注意: 同一股票可能多次定增，因此不能按股票代码映射，
    必须按行序逐行写入（result_df 与 Excel 行数一致、顺序一致）。
    """
    import openpyxl

    target = output_path or scored_excel_path
    wb = openpyxl.load_workbook(target)
    ws = wb.active

    # 如果已有旧列则先删除（幂等写入）
    header_row = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    for col_name in ['盈利概率_LightGBM', '盈利概率_逻辑回归']:
        if col_name in header_row:
            col_idx = header_row.index(col_name) + 1
            ws.delete_cols(col_idx)
            # 重新读取 header（删除列后索引变化）
            header_row = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]

    # 写入两列
    lgb_col = ws.max_column + 1
    lr_col = lgb_col + 1
    ws.cell(1, lgb_col, '盈利概率_LightGBM')
    ws.cell(1, lr_col, '盈利概率_逻辑回归')

    # 按行序写入（row 2 = 第1条数据, 与 result_df 第0行对应）
    matched = 0
    data_rows = ws.max_row - 1  # 排除 header
    n = min(len(result_df), data_rows)
    for i in range(n):
        r = i + 2  # Excel 行号
        ws.cell(r, lgb_col, f'{result_df.iloc[i]["盈利概率_LightGBM"]*100:.1f}%')
        ws.cell(r, lr_col, f'{result_df.iloc[i]["盈利概率_逻辑回归"]*100:.1f}%')
        matched += 1

    wb.save(target)
    print(f'  ✅ 写入 Excel: {matched} 条 → {target}')
    return matched


def main():
    import argparse
    parser = argparse.ArgumentParser(description='定增盈利概率预测（双模型·全量特征）')
    parser.add_argument('scored_excel', help='已评分 Excel 路径')
    parser.add_argument('--output', default=None, help='输出路径（默认覆盖原文件）')
    args = parser.parse_args()

    print('=' * 60)
    print('定增盈利概率预测（LightGBM + 逻辑回归·全量特征）')
    print('=' * 60)

    result_df = predict(args.scored_excel)

    # 汇总统计
    print('\n  汇总:')
    print(f'    LightGBM 均值: {result_df["盈利概率_LightGBM"].mean()*100:.1f}%')
    print(f'    逻辑回归 均值: {result_df["盈利概率_逻辑回归"].mean()*100:.1f}%')
    high_lgb = (result_df['盈利概率_LightGBM'] > 0.5).sum()
    high_lr = (result_df['盈利概率_逻辑回归'] > 0.5).sum()
    print(f'    LightGBM >50%: {high_lgb} ({high_lgb/len(result_df)*100:.1f}%)')
    print(f'    逻辑回归 >50%: {high_lr} ({high_lr/len(result_df)*100:.1f}%)')

    # 写入 Excel
    write_to_excel(args.scored_excel, result_df, args.output)

    print('\n' + '=' * 60)
    print('🎉 预测完成!')
    print('=' * 60)


if __name__ == '__main__':
    main()
