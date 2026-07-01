#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正确计算100分制评分卡
"""

import numpy as np
import pandas as pd

# 读取现有评分卡数据
scorecard_df = pd.read_csv('ml_training/output/scorecard_table.csv')

print("=== 100分制评分卡正确计算 ===")
print()

# 标准评分卡公式：
# Score = Offset + Factor × ln(odds)
# 其中：Factor = PDO / ln(2)
#       Offset = Base_Score - Factor × ln(Base_Odds)

# 设定100分制参数
PDO = 20  # Points to Double the Odds (标准值)
base_score_100 = 100  # 在50%概率时的目标分数
base_odds = 1  # 基准几率（1:1，即50%概率）

Factor = PDO / np.log(2)  # ≈ 28.8539
Offset = base_score_100 - (Factor * np.log(base_odds))  # = 100

print(f"100分制标准参数:")
print(f"  PDO (几率翻倍分数): {PDO}")
print(f"  基准分: {base_score_100}分 (在50%概率时)")
print(f"  基准几率: {base_odds}:1")
print()

print(f"计算结果:")
print(f"  Factor = PDO / ln(2) = {PDO} / {np.log(2):.4f} = {Factor:.4f}")
print(f"  Offset = {base_score_100} - ({Factor:.4f} × {np.log(base_odds):.4f}) = {Offset:.4f}")
print()

print(f"100分制标准公式:")
print(f"  Score = {Offset:.4f} + {Factor:.4f} × ln(odds)")
print(f"  即：Score = 100 + {Factor:.4f} × ln(odds)")
print()

print("=== 重新计算各特征分箱得分 ===")
print()

# 从原始评分卡中提取数据
features_df = scorecard_df[scorecard_df['特征'] != '基础分'].copy()
features_df['系数'] = pd.to_numeric(features_df['系数'], errors='coerce')
features_df['WOE'] = pd.to_numeric(features_df['WOE'], errors='coerce')

# 计算每个特征的100分制得分
scorecard_100pt = []

for feature in features_df['特征'].unique():
    feat_data = features_df[features_df['特征'] == feature].copy()
    coef = feat_data['系数'].iloc[0]

    print(f"特征: {feature}")
    print(f"  逻辑回归系数: {coef:.4f}")
    print(f"  评分卡系数: {Factor:.4f} × {coef:.4f} = {Factor * coef:.4f}")
    print(f"  分箱范围 | WOE值 | 100分制得分")
    print("-" * 50)

    for _, row in feat_data.iterrows():
        bin_range = row['分箱范围']
        woe = pd.to_numeric(row['WOE'], errors='coerce')

        # 100分制得分 = Factor × 系数 × WOE
        new_score = Factor * coef * woe

        print(f"{bin_range:20} | {woe:7.4f} | {new_score:7.2f}")

        scorecard_100pt.append({
            '特征': feature,
            '分箱范围': bin_range,
            'WOE值': woe,
            '100分制得分': round(new_score, 2),
            '逻辑回归系数': coef,
            '评分卡系数': Factor * coef
        })

    print()

# 保存100分制评分卡
scorecard_100pt_df = pd.DataFrame(scorecard_100pt)
scorecard_100pt_df.to_csv('ml_training/output/scorecard_correct_100pt.csv', index=False, encoding='utf-8-sig')
print(f"✅ 100分制评分卡已保存: ml_training/output/scorecard_correct_100pt.csv")

print()
print("=== 最终100分制评分卡公式 ===")
print(f"总分 = 100 + Σ(各特征100分制得分)")
print(f"其中每个特征得分 = {Factor:.4f} × 逻辑回归系数 × WOE值")
print()

# 计算得分范围
max_score_contribution = abs(scorecard_100pt_df['100分制得分']).max()
min_total = 100 - max_score_contribution * len(features_df['特征'].unique())
max_total = 100 + max_score_contribution * len(features_df['特征'].unique())

print(f"理论总分范围: {min_total:.1f} 到 {max_total:.1f}")
print(f"实际考虑: 不是所有特征都会同时达到极值，预期范围约 80-120分")
print()
print("=== 得分解释 ===")
print(f"  < 80分: 低盈利概率，谨慎参与")
print(f"  80-100分: 中等盈利概率，可考虑参与")
print(f"  100-120分: 高盈利概率，建议参与")
print(f"  > 120分: 极高盈利概率，重点关注")