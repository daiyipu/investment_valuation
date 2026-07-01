#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重新理解真正的100分制评分卡
"""

import numpy as np
import pandas as pd
from scipy import stats

print("=== 理解评分卡分数范围 ===")
print()

# 读取现有评分卡数据
scorecard_df = pd.read_csv('price_maintenance_risk_analysis/ml_training/output/scorecard_table.csv')
features_df = scorecard_df[scorecard_df['特征'] != '基础分'].copy()
features_df['得分'] = pd.to_numeric(features_df['得分'], errors='coerce')

# 现有600分制的分析
base_score_original = 591.3
feature_scores = features_df['得分'].values

print("现有600分制评分卡分析:")
print(f"  基础分: {base_score_original}")
print(f"  特征得分范围: {feature_scores.min():.1f} 到 {feature_scores.max():.1f}")
print(f"  理论总分范围: {base_score_original + feature_scores.min():.1f} 到 {base_score_original + feature_scores.max():.1f}")

# 但实际上，一个样本不会所有特征都取极值
# 假设18个特征，每个特征期望贡献接近0（WOE值有正有负）
expected_total = base_score_original
print(f"  期望总分: {expected_total:.1f} (基础分)")
print()

# 标准评分卡的工作原理
print("标准评分卡原理:")
print("  1. 设定在某个概率时，分数为基准分")
print("  2. PDO表示几率翻倍需要的分数变化")
print("  3. 分数可以超过基准分，因为概率可能很高")
print("  4. 分数也可以低于基准分，因为概率可能很低")
print()

# 比如FICO分数: 300-850范围，不是说"850分制"，而是说在这个范围内
print("类比: FICO信用分数范围300-850")
print("     不是说'850分制'，而是分数分布在这个范围内")
print()

# 如果要真正的100分制（0-100范围），需要重新设计
print("=== 真正的100分制设计 ===")
print()

# 选项1: 将现有600分制映射到0-100
min_original = base_score_original + feature_scores.min() * 18  # 极端情况
max_original = base_score_original + feature_scores.max() * 18  # 极端情况

# 但更合理的是看实际分布
# 假设实际分布大概在550-650之间
practical_min = 550
practical_max = 650

print(f"选项1: 线性映射现有分数到0-100")
print(f"  现有实际范围: {practical_min}-{practical_max}")
print(f"  映射公式: Score_100 = (Score_原始 - {practical_min}) / {practical_max - practical_min} × 100")
print(f"  即: Score_100 = (Score_原始 - {practical_min}) / {practical_max - practical_min} × 100")
print()

# 映射参数
scale_factor = 100 / (practical_max - practical_min)
offset = -practical_min * scale_factor

print(f"  简化: Score_100 = {offset:.2f} + {scale_factor:.4f} × Score_原始")
print()

# 验证映射
test_cases = [550, 575, 600, 625, 650]
print("  验证映射:")
for score in test_cases:
    mapped = offset + scale_factor * score
    print(f"    {score}分 → {mapped:.1f}分")

print()

# 选项2: 重新设计评分卡，强制总分在0-100
print("选项2: 重新设计评分卡，确保总分在0-100")
print("  需要调整基础分和PDO，使得:")
print("  - 最低概率时接近0分")
print("  - 最高概率时接近100分")
print("  - 50%概率时为50分")
print()

# 从训练数据中统计实际的概率分布
# 这里我们假设盈利概率大概在20%-80%之间
# 对应的logit范围:
prob_low = 0.2
prob_high = 0.8
logit_low = np.log(prob_low / (1 - prob_low))  # 约-1.39
logit_high = np.log(prob_high / (1 - prob_high))  # 约+1.39

print(f"  假设实际概率范围: {prob_low*100:.0f}%-{prob_high*100:.0f}%")
print(f"  对应logit范围: {logit_low:.2f} 到 {logit_high:.2f}")
print()

# 重新计算评分卡参数
# 目标: 概率20%时0分，概率80%时100分
# 0 = Offset + Factor × logit_low
# 100 = Offset + Factor × logit_high
# 解出: Factor = 100 / (logit_high - logit_low) = 100 / (1.39 - (-1.39)) ≈ 36
#      Offset = -Factor × logit_low

Factor_new = 100 / (logit_high - logit_low)
Offset_new = -Factor_new * logit_low

print(f"  新参数:")
print(f"    Factor = {Factor_new:.2f}")
print(f"    Offset = {Offset_new:.2f}")
print(f"  验证:")
print(f"    {prob_low*100:.0f}%概率: {Offset_new + Factor_new * logit_low:.1f}分")
print(f"    {prob_high*100:.0f}%概率: {Offset_new + Factor_new * logit_high:.1f}分")
print(f"    50%概率: {Offset_new + Factor_new * 0:.1f}分")
print()

print("结论:")
print("  标准评分卡的'100分制'是指在某个基准条件下为100分")
print("  不是说总分被限制在100分以内")
print("  如果要严格的0-100分制，需要重新设计参数")