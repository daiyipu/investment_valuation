# -*- coding: utf-8 -*-
"""
收益率转换工具使用示例

本文件展示如何在实际项目中使用收益率转换工具。

作者: 投资估值系统项目组
日期: 2026-04-01
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import numpy as np
import json
from utils.return_conversion import (
    discrete_to_continuous,
    continuous_to_discrete,
    annualize_discrete_return,
    annualize_continuous_return,
    get_gbm_drift_from_discrete,
    d2c,
    c2d
)

print("=" * 80)
print("收益率转换工具使用示例")
print("=" * 80)

# ============================================================================
# 示例1: 基础转换
# ============================================================================
print("\n" + "=" * 80)
print("【示例1】基础转换：离散复利 ↔ 连续复利")
print("=" * 80)

# 离散复利（报告中的数据）
discrete_return = -0.176  # -17.60%
print(f"\n离散复利（报告显示）: {discrete_return*100:+.2f}%")

# 转换为连续复利（用于GBM模型）
continuous_return = discrete_to_continuous(discrete_return)
print(f"连续复利（GBM模型）: {continuous_return*100:+.2f}%")
print(f"差异: {abs(continuous_return - discrete_return)*100:.2f}%")

# 反向转换
discrete_back = continuous_to_discrete(continuous_return)
print(f"反向转换验证: {discrete_back*100:+.2f}%")
print(f"转换误差: {abs(discrete_return - discrete_back):.6f}")

# ============================================================================
# 示例2: 年化收益率对比
# ============================================================================
print("\n" + "=" * 80)
print("【示例2】年化收益率对比：离散 vs 连续")
print("=" * 80)

period_return = -0.176  # 20日离散收益率
window = 20

# 离散复利年化（报告展示用）
annual_discrete = annualize_discrete_return(period_return, window)
print(f"\n区间离散收益率: {period_return*100:+.2f}%")
print(f"离散年化（报告用）: {annual_discrete*100:+.2f}%")

# 连续复利年化（GBM模型用）
annual_continuous = get_gbm_drift_from_discrete(period_return, window)
print(f"连续年化（GBM用）: {annual_continuous*100:+.2f}%")
print(f"差异: {abs(annual_continuous - annual_discrete)*100:.2f}%")

# ============================================================================
# 示例3: 从市场数据文件读取并转换
# ============================================================================
print("\n" + "=" * 80)
print("【示例3】从市场数据文件读取并转换")
print("=" * 80)

# 读取市场数据
data_file = '../data/300735_SZ_market_data.json'
if os.path.exists(data_file):
    with open(data_file, 'r') as f:
        market_data = json.load(f)

    print(f"\n股票: {market_data['stock_name']} ({market_data['stock_code']})")

    # 报告中的数据（离散复利）
    print(f"\n【报告中的数据（离散复利）】")
    print(f"  20日区间收益率: {market_data['period_return_20d']*100:+.2f}%")
    print(f"  60日区间收益率: {market_data['period_return_60d']*100:+.2f}%")
    print(f"  20日年化收益率: {market_data['annual_return_20d']*100:+.2f}%")
    print(f"  60日年化收益率: {market_data['annual_return_60d']*100:+.2f}%")

    # 转换为连续复利（GBM模型用）
    print(f"\n【转换为连续复利（GBM模型用）】")

    # 20日窗口
    drift_20d = get_gbm_drift_from_discrete(market_data['period_return_20d'], 20)
    print(f"  20日GBM漂移率: {drift_20d*100:+.2f}%")
    print(f"    计算: ln(1 + {market_data['period_return_20d']:.4f}) × (250/20)")
    print(f"         = {discrete_to_continuous(market_data['period_return_20d']):.4f} × 12.5")
    print(f"         = {drift_20d:.4f}")

    # 60日窗口
    drift_60d = get_gbm_drift_from_discrete(market_data['period_return_60d'], 60)
    print(f"  60日GBM漂移率: {drift_60d*100:+.2f}%")
    print(f"    计算: ln(1 + {market_data['period_return_60d']:.4f}) × (250/60)")
    print(f"         = {discrete_to_continuous(market_data['period_return_60d']):.4f} × 4.167")
    print(f"         = {drift_60d:.4f}")

    # 波动率（已经是连续复利）
    print(f"\n【波动率（已经是连续复利）】")
    print(f"  20日波动率: {market_data['volatility_20d']*100:.2f}%")
    print(f"  60日波动率: {market_data['volatility_60d']*100:.2f}%")

    # ============================================================================
    # 示例4: 在蒙特卡洛模拟中使用
    # ============================================================================
    print(f"\n" + "=" * 80)
    print("【示例4】在蒙特卡洛模拟中使用")
    print("=" * 80)

    from utils.analysis_tools import PrivatePlacementRiskAnalyzer

    # 创建定增项目分析器
    analyzer = PrivatePlacementRiskAnalyzer(
        issue_price=25.80,
        issue_shares=1000000,
        lockup_period=6,
        current_price=market_data['current_price']
    )

    # 从市场数据获取参数（离散复利）
    period_return = market_data['period_return_60d']  # -14.62%
    volatility = market_data['volatility_60d']  # 33.80%

    # 转换为连续复利（GBM漂移率）
    drift = get_gbm_drift_from_discrete(period_return, 60)

    print(f"\n蒙特卡洛模拟参数:")
    print(f"  当前价格: {market_data['current_price']:.2f} 元")
    print(f"  漂移率（连续复利年化）: {drift*100:+.2f}%")
    print(f"  波动率（连续复利年化）: {volatility*100:.2f}%")
    print(f"  锁定期: 6个月")

    # 运行蒙特卡洛模拟
    print(f"\n运行蒙特卡洛模拟...")
    simulations = analyzer.monte_carlo_simulation(
        n_simulations=1000,
        time_steps=180,  # 6个月约180天
        drift=drift,
        volatility=volatility,
        seed=42
    )

    # 分析结果
    final_prices = simulations[f'Day_{180}']
    profit_prob = (final_prices > 25.80).mean() * 100
    mean_price = final_prices.mean()
    median_price = final_prices.median()

    print(f"\n模拟结果（1000次）:")
    print(f"  盈利概率（>发行价）: {profit_prob:.2f}%")
    print(f"  期望价格: {mean_price:.2f} 元")
    print(f"  中位数价格: {median_price:.2f} 元")
    print(f"  95%置信区间: [{final_prices.quantile(0.025):.2f}, {final_prices.quantile(0.975):.2f}]")

    # ============================================================================
    # 示例5: 快捷函数使用
    # ============================================================================
    print(f"\n" + "=" * 80)
    print("【示例5】快捷函数使用")
    print("=" * 80)

    # 使用快捷别名
    discrete = -0.20
    continuous_fast = d2c(discrete)  # 等同于 discrete_to_continuous
    back_to_discrete = c2d(continuous_fast)  # 等同于 continuous_to_discrete

    print(f"\n使用快捷函数:")
    print(f"  d2c({discrete}) = {continuous_fast:.6f}")
    print(f"  c2d({continuous_fast:.6f}) = {back_to_discrete:.6f}")

else:
    print(f"\n⚠️  市场数据文件不存在: {data_file}")
    print(f"   请先运行 update_market_data.py 生成数据")

# ============================================================================
# 总结
# ============================================================================
print("\n" + "=" * 80)
print("【使用总结】")
print("=" * 80)

print("""
📋 收益率类型选择指南：

1. 【报告展示】使用离散复利
   - 公式: (P_end - P_start) / P_start
   - 优点: 直观易懂
   - 函数: calculate_period_return() / calculate_annual_return()

2. 【GBM模型】使用连续复利
   - 公式: ln(P_end / P_start)
   - 优点: 数学性质好，符合伊藤引理
   - 转换: get_gbm_drift_from_discrete()

3. 【典型使用场景】
   报告数据（离散）→ 转换 → GBM模型（连续）

   示例代码:
   # 步骤1: 获取报告数据（离散复利）
   period_return = market_data['period_return_60d']  # -14.62%
   volatility = market_data['volatility_60d']        # 33.80%

   # 步骤2: 转换为连续复利（GBM模型用）
   drift = get_gbm_drift_from_discrete(period_return, 60)

   # 步骤3: 用于GBM模型
   monte_carlo_simulation(drift=drift, volatility=volatility)
""")

print("=" * 80)
print("✓ 示例运行完成")
print("=" * 80)
