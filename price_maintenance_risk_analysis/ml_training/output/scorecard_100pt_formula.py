#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将评分卡转换为100分制标准形式
"""

import pandas as pd
import numpy as np

# 读取现有评分卡
scorecard_df = pd.read_csv('ml_training/output/scorecard_table.csv')

print("=== 现有评分卡分析 ===")
print(f"基础分: 591.3分")
print(f"评分公式: 总分 = 591.3 + Σ(各特征分箱得分)")
print()

# 提取基础分
base_score = 591.3
B = 20 / np.log(2)  # ≈ 28.85

# 分析得分范围
non_base = scorecard_df[scorecard_df['特征'] != '基础分']
min_score = non_base['得分'].min()
max_score = non_base['得分'].max()
score_range = max_score - min_score

print(f"单特征分箱得分范围: {min_score:.1f} 到 {max_score:.1f}")
print(f"得分跨度: {score_range:.1f}分")
print()

# 计算转换到100分制的参数
# 假设原始总分范围约550-650（从得分分布可以看出）
# 我们要映射到0-100分制

# 方法1: 基于WOE值直接计算logit得分
# logit = intercept + Σ(coef_i × woe_i)
# 然后将logit映射到0-100

print("=== 转换到100分制标准形式 ===")
print()

# 提取特征和WOE信息
features = non_base['特征'].unique()
print(f"特征数量: {len(features)}")
print()

# 为每个特征构建分箱得分表
print("=== 100分制评分卡公式 ===")
print("得分 = 截距 + Σ(斜率i × 特征i的分箱WOE值)")
print()

# 计算归一化参数
# 从现有得分推算：
# 现有: 得分 = B × coef × woe
# 目标: 得分 = new_coef × woe (归一化到100分制)

# 重新计算归一化系数
normalization_factor = 100 / (18 * score_range)  # 粗略归一化
print(f"归一化因子: {normalization_factor:.4f}")
print()

print("=== 各特征的100分制分箱得分 ===")
scorecard_100pt = []

for feature in features:
    feat_data = non_base[non_base['特征'] == feature].copy()

    # 提取该特征的系数，确保是数值类型
    coef = pd.to_numeric(feat_data['系数'].iloc[0], errors='coerce')

    print(f"\\n特征: {feature}")
    print(f"逻辑回归系数: {coef:.4f}")
    print(f"分箱范围 | WOE值 | 现有得分 | 100分制得分")
    print("-" * 50)

    for _, row in feat_data.iterrows():
        bin_range = row['分箱范围']
        woe = pd.to_numeric(row['WOE'], errors='coerce')
        old_score = pd.to_numeric(row['得分'], errors='coerce')

        # 100分制得分 = 现有得分归一化
        # 简单方法: 按比例缩放
        new_score = (old_score / score_range) * 20  # 缩放到合理范围

        print(f"{bin_range:20} | {woe:7.4f} | {old_score:7.1f} | {new_score:7.2f}")

        scorecard_100pt.append({
            '特征': feature,
            '分箱范围': bin_range,
            'WOE值': woe,
            '现有得分': old_score,
            '100分制得分': round(new_score, 2),
            '系数': coef
        })

# 保存100分制评分卡
scorecard_100pt_df = pd.DataFrame(scorecard_100pt)
scorecard_100pt_df.to_csv('ml_training/output/scorecard_100pt.csv', index=False, encoding='utf-8-sig')
print(f"\\n100分制评分卡已保存: ml_training/output/scorecard_100pt.csv")

print("\\n=== 简化的评分公式 ===")
print("总分 = 50 + Σ(各特征100分制得分)")
print("说明: 基础分50分，每个特征根据分箱贡献不同得分，总分接近100分")