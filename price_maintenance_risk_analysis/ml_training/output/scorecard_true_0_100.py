#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
真正的0-100分制评分卡设计
"""

import numpy as np
import pandas as pd

# 读取现有评分卡数据
scorecard_df = pd.read_csv('price_maintenance_risk_analysis/ml_training/output/scorecard_table.csv')
features_df = scorecard_df[scorecard_df['特征'] != '基础分'].copy()
features_df['系数'] = pd.to_numeric(features_df['系数'], errors='coerce')
features_df['WOE'] = pd.to_numeric(features_df['WOE'], errors='coerce')

print("=== 真正的0-100分制评分卡设计 ===")
print()

# 假设实际概率分布范围（基于训练数据）
# 定增盈利概率大概在10%-90%之间
prob_min = 0.10  # 最低盈利概率10%
prob_max = 0.90  # 最高盈利概率90%
prob_base = 0.50  # 基准概率50%

# 对应的logit值
logit_min = np.log(prob_min / (1 - prob_min))  # 约-2.20
logit_max = np.log(prob_max / (1 - prob_max))  # 约+2.20
logit_base = np.log(prob_base / (1 - prob_base))  # = 0

print(f"概率范围设计:")
print(f"  最低概率: {prob_min*100:.0f}% → logit = {logit_min:.2f} → 目标分数: 0分")
print(f"  基准概率: {prob_base*100:.0f}% → logit = {logit_base:.2f} → 目标分数: 50分")
print(f"  最高概率: {prob_max*100:.0f}% → logit = {logit_max:.2f} → 目标分数: 100分")
print()

# 计算新的评分卡参数
# 目标方程:
# 0 = Offset + Factor × logit_min
# 50 = Offset + Factor × logit_base
# 100 = Offset + Factor × logit_max

# 从前两个方程解出:
# Factor = (50 - 0) / (logit_base - logit_min) = 50 / (0 - logit_min)
Factor_new = 50 / (logit_base - logit_min)
Offset_new = 0 - Factor_new * logit_min

print(f"新评分卡参数:")
print(f"  Factor = {Factor_new:.2f}")
print(f"  Offset = {Offset_new:.2f}")
print()

# 验证
print("验证:")
print(f"  {prob_min*100:.0f}%概率: {Offset_new + Factor_new * logit_min:.1f}分 (应为0)")
print(f"  {prob_base*100:.0f}%概率: {Offset_new + Factor_new * logit_base:.1f}分 (应为50)")
print(f"  {prob_max*100:.0f}%概率: {Offset_new + Factor_new * logit_max:.1f}分 (应为100)")
print()

# 对应的PDO
# PDO = ln(2) × Factor
PDO_new = np.log(2) * Factor_new
print(f"对应的PDO: {PDO_new:.2f}")
print(f"(几率翻倍需要{PDO_new:.1f}分)")
print()

print("=== 重新计算各特征的0-100分制得分 ===")
print()

# 重新计算各特征的分箱得分
scorecard_true_100 = []

for feature in features_df['特征'].unique():
    feat_data = features_df[features_df['特征'] == feature].copy()
    coef = feat_data['系数'].iloc[0]

    print(f"特征: {feature}")
    print(f"  逻辑回归系数: {coef:.4f}")
    print(f"  评分卡系数: {Factor_new:.2f} × {coef:.4f} = {Factor_new * coef:.4f}")

    for _, row in feat_data.iterrows():
        bin_range = row['分箱范围']
        woe = pd.to_numeric(row['WOE'], errors='coerce')

        # 新的得分 = Factor_new × 系数 × WOE
        new_score = Factor_new * coef * woe

        scorecard_true_100.append({
            '特征': feature,
            '分箱范围': bin_range,
            'WOE值': woe,
            '0-100分制得分': round(new_score, 2),
            '逻辑回归系数': coef,
            '评分卡系数': Factor_new * coef
        })

print()

# 保存新的评分卡
scorecard_true_100_df = pd.DataFrame(scorecard_true_100)
scorecard_true_100_df.to_csv('price_maintenance_risk_analysis/ml_training/output/scorecard_true_0_100.csv', index=False, encoding='utf-8-sig')
print("✅ 真正0-100分制评分卡已保存")

print()
print("=== 最终0-100分制评分卡公式 ===")
print(f"总分 = {Offset_new:.1f} + Σ(各特征分箱得分)")
print(f"其中每个特征得分 = {Factor_new:.2f} × 逻辑回归系数 × WOE值")
print()

# 分析得分范围
max_contribution = abs(scorecard_true_100_df['0-100分制得分']).max()
num_features = len(features_df['特征'].unique())

print(f"分数范围分析:")
print(f"  单特征最大贡献: ±{max_contribution:.2f}分")
print(f"  特征总数: {num_features}个")
print(f"  理论极值范围: {Offset_new:.1f} ± {max_contribution * num_features:.1f} ≈ {Offset_new - max_contribution * num_features:.1f} 到 {Offset_new + max_contribution * num_features:.1f}")
print(f"  实际合理范围: 0-100分 (对应概率{prob_min*100:.0f}%-{prob_max*100:.0f}%)")
print()

print("得分解释:")
print(f"  0-30分: 低盈利概率 ({prob_min*100:.0f}%-30%)，谨慎参与")
print(f"  30-50分: 中低盈利概率 (30%-50%)，观察为主")
print(f"  50-70分: 中等盈利概率 (50%-70%)，可考虑参与")
print(f"  70-90分: 高盈利概率 (70%-{prob_max*100:.0f}%)，建议参与")
print(f"  90-100分: 极高盈利概率 (>85%)，重点关注")