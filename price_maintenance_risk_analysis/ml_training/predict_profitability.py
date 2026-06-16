#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增盈利概率预测 - 加载训练好的模型对新样本预测

模型版本由 model_registry 的 current.full / current.scorecard 指针决定
（见 manage_models.py current）。不再写死具体版本/AUC。

支持两种用法:
  1. 独立运行: python ml_training/predict_profitability.py <scored_excel> [--output result.xlsx]
  2. 被调用:   from predict_profitability import predict; df = predict(scored_excel_path)

输出三列:
  - 盈利概率_LightGBM
  - 盈利概率_逻辑回归
  - 评分卡得分（base=600，PDO=20；scorecard 未注册时省略）
"""

import sys
import os
import re
import json
import warnings
import pickle
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)


_INTERVAL_RE = re.compile(r'^[\[(](.*?),\s*(.*?)[\])]$')


def _parse_interval(key):
    """解析 pandas 区间字符串 '(0.33, 2.145]' → (0.33, 2.145)。失败返回 None。"""
    m = _INTERVAL_RE.match(str(key).strip())
    if not m:
        return None
    try:
        return float(m.group(1)), float(m.group(2))
    except ValueError:
        return None


def score_with_scorecard(df, sc_dir):
    """用评分卡模型对 df 每行打分。

    score = base_points + Σ(B · coef_i · woe_i)
    缺失/越界特征按最近 bin 兜底；NaN 贡献 0（与训练时 evaluate_scorecard 的
    fillna(0) 行为一致）。

    Args:
        df: 已构造完整特征的 DataFrame
        sc_dir: 评分卡版本目录（含 scorecard_model.pkl）

    Returns:
        np.array[float] 每行得分
    """
    with open(os.path.join(sc_dir, 'scorecard_model.pkl'), 'rb') as f:
        sc = pickle.load(f)
    model = sc['model']
    features = sc['features']
    woe_bins = sc['woe_bins']
    base_points = float(sc['scoring_params']['base_points'])
    B = float(sc['scoring_params']['B'])
    coefs = dict(zip(features, model.coef_[0]))

    scores = np.full(len(df), base_points, dtype=float)
    for feat in features:
        if feat not in df.columns or feat not in woe_bins:
            continue
        # 预解析该特征的 bins: [(left, right, woe), ...] 按 left 排序
        parsed = []
        for k, woe in woe_bins[feat].get('woe_map', {}).items():
            iv = _parse_interval(k)
            if iv is not None:
                parsed.append((iv[0], iv[1], float(woe)))
        if not parsed:
            continue
        parsed.sort(key=lambda x: x[0])
        min_left = parsed[0][0]
        first_woe = parsed[0][2]
        last_woe = parsed[-1][2]

        vals = pd.to_numeric(df[feat], errors='coerce').values
        coef = coefs[feat]
        for i, v in enumerate(vals):
            if v != v:  # NaN → 跳过（贡献 0）
                continue
            woe = None
            for l, r, w in parsed:
                if l < v <= r:
                    woe = w
                    break
            if woe is None:
                woe = first_woe if v <= min_left else last_woe  # 越界兜底
            scores[i] += B * coef * woe
    return scores


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
    from export_features import load_db_features, load_scored_features, load_financial_ratios
    from derive_features import (
        derive_fcf_growth_rates, derive_fcf_cross_metrics,
        derive_financial_score_deltas, derive_valuation_relative,
        derive_market_momentum, derive_industry_valuation_growth,
        derive_market_index_features,
    )
    from model_registry import require_current_dir, get_current
    from db_model_store import load_predict_bundle   # full 模型权重+meta 从 DB 加载(不再读磁盘 version 目录)

    output_dir = os.path.join(SCRIPT_DIR, 'output')
    # 模型从 registry 的当前生产版本读取（见 manage_models.py current）
    version = get_current('full')
    bundle = load_predict_bundle(version)   # 权重+features+medians 从 DB ml_model_meta 加载
    print(f'  [ML-0] 使用 full 模型版本: {version} (权重从 DB 加载)')
    # 评分卡（可选：未注册则跳过评分卡得分列）
    sc_version = get_current('scorecard')
    sc_dir = None
    if sc_version:
        try:
            sc_dir = require_current_dir('scorecard')
            print(f'  [ML-0] 使用 scorecard 模型版本: {sc_version}')
        except RuntimeError as e:
            print(f'  [ML-0] 评分卡版本目录缺失，跳过评分卡得分: {e}')

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
        issue_date = str(row.get('报价日', '')).replace('.0', '').strip()
        if not issue_date or issue_date == 'nan' or len(issue_date) < 8:
            issue_date = None
        sample_keys.append((code, issue_date))

    db_feats = load_db_features(sample_keys)
    matched_price = db_feats.get('当前价', pd.Series()).notna().sum()
    print(f'    行情匹配: {matched_price}/{len(scored)}')

    # 财务比率（financial_indicators API；三表已弃用，覆盖率0.43%且与比率表重复）
    try:
        ratio_feats = load_financial_ratios(sample_keys)
    except Exception as e:
        print(f'    财务比率跳过: {e}')
        ratio_feats = pd.DataFrame()

    # ═══════════════════════════════════════════════
    # 3. 合并
    # ═══════════════════════════════════════════════
    scored = scored.reset_index(drop=True)
    # 去掉 db_feats 中与 scored 同名的列（报价日/定增决策/报价日价格…），
    # 一律取 scored（权威源），避免 concat 后出现重名列导致 df[c] 变成多列。
    _dup_cols = (set(scored.columns) & set(db_feats.columns)) - {'股票代码'}
    if _dup_cols:
        print(f'    去重列(取scored版): {sorted(_dup_cols)}')
    db_feat_cols = [c for c in db_feats.columns if c != '股票代码' and c not in _dup_cols]
    df = pd.concat([scored, db_feats[db_feat_cols].reset_index(drop=True)], axis=1)

    if not ratio_feats.empty:
        df = pd.concat([df.reset_index(drop=True), ratio_feats.reset_index(drop=True)], axis=1)

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
    print('  [ML-4] 加载模型 meta(从 DB)...')
    model_features = bundle['features']
    train_medians = bundle['medians']
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
    # 7/8. 预测(LGB+LR 或 评分卡 SC 二选一)
    # ═══════════════════════════════════════════════
    is_sc = not bundle.get('lgb_model')   # 评分卡模型: lgb_model 为空
    if is_sc:
        # 评分卡路径: WOE 分箱 + LR 打分(替代 LGB+LR)
        print('  [ML-6/7] 评分卡(SC) 预测(WOE+LR)...')
        from eval_loyo import apply_woe
        sc = pickle.loads(bundle['lr_bundle'])   # {kind, woe_bins, lr_model, features, medians}
        X_woe = apply_woe(X_pred, sc['features'], sc['woe_bins']).replace([np.inf, -np.inf], 0).fillna(0)
        lgb_proba = sc['lr_model'].predict_proba(X_woe)[:, 1]
        lr_proba = lgb_proba                      # SC 即主分数, 无独立 LR
        print(f'    SC 概率范围: [{lgb_proba.min():.3f}, {lgb_proba.max():.3f}]')
    else:
        print('  [ML-6] LightGBM 预测...')
        booster = lgb.Booster(model_str=bundle['lgb_model'])
        lgb_feature_names = booster.feature_name()
        X_lgb = pd.DataFrame(index=X_pred.index)
        for feat in lgb_feature_names:
            X_lgb[feat] = X_pred[feat] if feat in X_pred.columns else 0
        X_lgb = X_lgb.astype(float).values
        lgb_proba = booster.predict(X_lgb)
        print(f'    概率范围: [{lgb_proba.min():.3f}, {lgb_proba.max():.3f}]')

        print('  [ML-7] 逻辑回归 预测...')
        lr_data = pickle.loads(bundle['lr_bundle'])
        lr_model, lr_scaler, lr_feature_names = lr_data['model'], lr_data['scaler'], lr_data['features']
        X_lr = pd.DataFrame(index=X_pred.index)
        for feat in lr_feature_names:
            X_lr[feat] = X_pred[feat] if feat in X_pred.columns else 0
        X_lr = X_lr.replace([np.inf, -np.inf], 0).fillna(0)
        lr_proba = lr_model.predict_proba(lr_scaler.transform(X_lr.values))[:, 1]
        print(f'    概率范围: [{lr_proba.min():.3f}, {lr_proba.max():.3f}]')

    # ═══════════════════════════════════════════════
    # 9. 评分卡得分（可选）
    # ═══════════════════════════════════════════════
    sc_scores = None
    if sc_dir is not None:
        print('  [ML-8] 评分卡得分...')
        try:
            sc_scores = score_with_scorecard(df, sc_dir)
            print(f'    得分范围: [{sc_scores.min():.0f}, {sc_scores.max():.0f}]'
                  f'  均值: {sc_scores.mean():.0f}')
        except Exception as e:
            print(f'    ⚠ 评分卡打分失败，跳过: {e}')
            sc_scores = None

    # ═══════════════════════════════════════════════
    # 10. 返回
    # ═══════════════════════════════════════════════
    result = pd.DataFrame({
        '股票代码': df['股票代码'],
        '盈利概率_LightGBM': lgb_proba,
        '盈利概率_逻辑回归': lr_proba,
    })
    if sc_scores is not None:
        result['评分卡得分'] = np.round(sc_scores).astype(int)

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

    # 待写入列（评分卡得分仅在 result_df 中存在时写）
    out_cols = ['盈利概率_LightGBM', '盈利概率_逻辑回归']
    if '评分卡得分' in result_df.columns:
        out_cols.append('评分卡得分')

    # 如果已有旧列则先删除（幂等写入）
    header_row = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    for col_name in list(out_cols):
        if col_name in header_row:
            col_idx = header_row.index(col_name) + 1
            ws.delete_cols(col_idx)
            header_row = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]

    # 依次写入列
    col_map = {}
    for col_name in out_cols:
        c = ws.max_column + 1
        ws.cell(1, c, col_name)
        col_map[col_name] = c

    # 按行序写入（row 2 = 第1条数据, 与 result_df 第0行对应）
    matched = 0
    data_rows = ws.max_row - 1  # 排除 header
    n = min(len(result_df), data_rows)
    for i in range(n):
        r = i + 2  # Excel 行号
        ws.cell(r, col_map['盈利概率_LightGBM'],
                f'{result_df.iloc[i]["盈利概率_LightGBM"]*100:.1f}%')
        ws.cell(r, col_map['盈利概率_逻辑回归'],
                f'{result_df.iloc[i]["盈利概率_逻辑回归"]*100:.1f}%')
        if '评分卡得分' in col_map:
            ws.cell(r, col_map['评分卡得分'], int(result_df.iloc[i]['评分卡得分']))
        matched += 1

    wb.save(target)
    print(f'  ✅ 写入 Excel: {matched} 条 → {target} (列: {", ".join(out_cols)})')
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
    if '评分卡得分' in result_df.columns:
        s = result_df['评分卡得分']
        print(f'    评分卡得分: 均值 {s.mean():.0f}, 区间 [{s.min()}, {s.max()}], >600 占比 '
              f'{(s>600).sum()/len(s)*100:.1f}%')

    # 写入 Excel
    write_to_excel(args.scored_excel, result_df, args.output)

    print('\n' + '=' * 60)
    print('🎉 预测完成!')
    print('=' * 60)


if __name__ == '__main__':
    main()
