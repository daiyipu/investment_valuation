#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标准评分卡公式计算
"""

import numpy as np
import pandas as pd

# 读取现有评分卡数据
scorecard_df = pd.read_csv('ml_training/output/scorecard_table.csv')

print("=== 标准评分卡公式 ===")
print()

# 标准评分卡参数设置
# 通常设定：
# 1. 在特定几率下（odds=1，即50%概率），分数为某个基准分（比如600分）
# 2. PDO=20，意味着几率翻倍需要增加20分

PDO = 20  # Points to Double the Odds (标准值)
base_score_target = 600  # 目标基准分
base_odds = 1  # 基准几率（1:1，即50%概率）

# 计算因子和偏移量
Factor = PDO / np.log(2)  # ≈ 28.85
Offset = base_score_target - (Factor * np.log(base_odds))  # ≈ 600

print(f"标准参数:")
print(f"  PDO (几率翻倍分数): {PDO}")
print(f"  基准分: {base_score_target}")
print(f"  基准几率: {base_odds}:1 (50%概率)")
print()

print(f"计算结果:")
print(f"  Factor = PDO / ln(2) = {PDO} / {np.log(2):.4f} = {Factor:.4f}")
print(f"  Offset = 基准分 - (Factor × ln(基准几率))")
print(f"         = {base_score_target} - ({Factor:.4f} × {np.log(base_odds):.4f})")
print(f"         = {Offset:.4f}")
print()

print("标准评分卡公式:")
print(f"  Score = {Offset:.4f} + {Factor:.4f} × ln(odds)")
print()

# 验证现有评分卡的计算
print("=== 验证现有评分卡 ===")

# 从scorecard_table.csv中提取数据
# 找到基础分行
base_row = scorecard_df[scorecard_df['特征'] == '基础分']
if not base_row.empty:
    existing_base_score = base_row['得分'].values[0]
    print(f"现有评分卡基础分: {existing_base_score:.1f}")

    # 推算对应的几率
    # 基础分 = Offset - Factor × intercept
    # intercept = (Offset - 基础分) / Factor
    inferred_intercept = (Offset - existing_base_score) / Factor
    print(f"推算的logistic回归截距: {inferred_intercept:.4f}")
    print()

    # 验证：在截距处（logits=0），几率=1，分数=基础分
    print("验证:")
    print(f"  当logits=0时，odds=1，ln(odds)=0")
    print(f"  Score = {Offset:.4f} + {Factor:.4f} × 0 = {Offset:.4f}")
    print(f"  与现有基础分{existing_base_score:.1f}的关系:")

    if abs(existing_base_score - Offset) < 1:
        print(f"  ✅ 完全一致！现有评分卡使用标准公式")
    else:
        diff = existing_base_score - Offset
        print(f"  ⚠️ 差异{diff:.1f}分，可能使用了不同的参数设置")

print()
print("=== 重新计算100分制评分卡 ===")
print()

# 如果我们要转换到100分制，需要重新设定基准分
# 比如：在50%概率时（odds=1），分数为50分
base_score_100 = 50
Factor_100 = Factor  # 保持相同的缩放因子
Offset_100 = base_score_100 - (Factor_100 * np.log(base_odds))

print(f"100分制参数:")
print(f"  基准分: {base_score_100}")
print(f"  基准几率: {base_odds}:1")
print(f"  Factor: {Factor_100:.4f} (保持不变)")
print(f"  Offset: {Offset_100:.4f}")
print()

print(f"100分制公式:")
print(f"  Score = {Offset_100:.4f} + {Factor_100:.4f} × ln(odds)")
print()

# 重新计算各特征的分箱得分
print("=== 重新计算100分制分箱得分 ===")

# 从原始评分卡中提取非基础分的特征
features_df = scorecard_df[scorecard_df['特征'] != '基础分'].copy()

# 确保数据类型正确
features_df['系数'] = pd.to_numeric(features_df['系数'], errors='coerce')
features_df['WOE'] = pd.to_numeric(features_df['WOE'], errors='coerce')
features_df['得分'] = pd.to_numeric(features_df['得分'], errors='coerce')

# 每个特征的得分 = Factor × 系数 × WOE
# 100分制得分与原始得分的差异只在于基础分的不同
# 因子Factor保持不变，所以特征得分保持不变
# 只是基础分从591.3调整为50

print("说明:")
print("  原始公式: Score_原始 = 591.3 + Σ(Factor × 系数 × WOE)")
print("  100分制: Score_100 = 50 + Σ(Factor × 系数 × WOE)")
print("  差异: Score_100 = Score_原始 - 541.3")
print()

# 验证计算
for feature in features_df['特征'].unique()[:3]:  # 只显示前3个特征
    feat_data = features_df[features_df['特征'] == feature]
    coef = feat_data['系数'].iloc[0]
    print(f"特征: {feature}")
    print(f"  系数: {coef:.4f}")
    print(f"  Factor × 系数 = {Factor:.4f} × {coef:.4f} = {Factor * coef:.4f}")

    # 显示第一个分箱的计算
    first_bin = feat_data.iloc[0]
    woe = first_bin['WOE']
    original_score = first_bin['得分']
    calculated_score = Factor * coef * woe

    print(f"  分箱 {first_bin['分箱范围']}:")
    print(f"    WOE: {woe:.4f}")
    print(f"    计算得分: {Factor:.4f} × {coef:.4f} × {woe:.4f} = {calculated_score:.1f}")
    print(f"    原始得分: {original_score:.1f}")

    if abs(calculated_score - original_score) < 0.1:
        print(f"    ✅ 验证通过")
    else:
        print(f"    ⚠️ 差异: {calculated_score - original_score:.1f}")
    print()

print("=== 100分制评分卡正确公式 ===")
print(f"总分 = 50 + Σ(各特征分箱得分)")
print(f"其中每个特征分箱得分 = {Factor:.4f} × 系数 × WOE值")
print(f"总分范围约: 50 ± {abs(features_df['得分']).max():.1f} ≈ {50 - abs(features_df['得分']).max():.1f} 到 {50 + abs(features_df['得分']).max():.1f}")