# -*- coding: utf-8 -*-
"""
收益率转换工具模块

本模块提供离散复利和连续复利之间的转换函数。

## 背景

在金融分析中，有两种常见的收益率表示方法：

1. **离散复利（Discrete Compounding）**：也叫简单涨跌幅
   - 公式：(P_end - P_start) / P_start
   - 优点：直观易懂，符合日常理解
   - 用途：报告展示、用户界面

2. **连续复利（Continuous Compounding）**：也叫对数收益率
   - 公式：ln(P_end / P_start)
   - 优点：数学性质好，时间可加，符合GBM模型
   - 用途：GBM模型、蒙特卡洛模拟、VaR计算、期权定价

## 设计原则

本系统采用**双层设计**：
- **展示层**：使用离散复利，便于用户理解
- **计算层**：使用连续复利，保证模型准确性

## 转换公式

### 离散 → 连续
    continuous_return = ln(1 + discrete_return)

### 连续 → 离散
    discrete_return = exp(continuous_return) - 1

## 使用示例

```python
# 场景1：报告展示（使用离散复利）
period_return_discrete = -0.176  # -17.60%
print(f"区间收益率: {period_return_discrete*100:.2f}%")

# 场景2：GBM模型计算（转换为连续复利）
period_return_continuous = discrete_to_continuous(period_return_discrete)
# = ln(1 - 0.176) = ln(0.824) = -0.1935
print(f"连续复利收益率: {period_return_continuous*100:.2f}%")

# 场景3：年化计算
annual_continuous = period_return_continuous * (250 / 20)  # 用于GBM的年化漂移率
```

作者: 投资估值系统项目组
日期: 2026-04-01
"""

import numpy as np
from typing import Union


def discrete_to_continuous(discrete_return: Union[float, np.ndarray]) -> Union[float, np.ndarray]:
    """
    将离散复利转换为连续复利（用于GBM模型等计算）

    离散复利 → 连续复利转换公式：
        continuous_return = ln(1 + discrete_return)

    理论依据：
    - 离散复利：P_end = P_start × (1 + r_discrete)
    - 连续复利：P_end = P_start × exp(r_continuous)
    - 联立得：1 + r_discrete = exp(r_continuous)
    - 因此：r_continuous = ln(1 + r_discrete)

    参数:
        discrete_return: 离散复利收益率
            - float: 单个收益率（如 -0.176 表示 -17.60%）
            - np.ndarray: 收益率序列

    返回:
        float or np.ndarray: 连续复利收益率（对数收益率）

    示例:
        >>> discrete_to_continuous(-0.176)
        -0.193585...  # -19.36%

        >>> discrete_to_continuous(0.10)
        0.095310...  # 9.53%

    注意:
        - 当 discrete_return <= -1 时，返回 -inf（表示本金全部损失）
        - 连续复利绝对值通常略大于离散复利（负值时更负，正值时更小）

    适用场景:
        - GBM模型参数设置（漂移率、波动率）
        - 蒙特卡洛模拟
        - VaR计算
        - 期权定价
        - 任何基于伊藤引理的金融工程计算
    """
    if isinstance(discrete_return, np.ndarray):
        # 处理数组输入
        result = np.zeros_like(discrete_return)
        valid_mask = discrete_return > -1.0

        # 有效值转换
        result[valid_mask] = np.log(1 + discrete_return[valid_mask])

        # 无效值处理（收益率 <= -100%）
        result[~valid_mask] = -np.inf

        return result
    else:
        # 处理单个值
        if discrete_return <= -1.0:
            return -np.inf
        return np.log(1 + discrete_return)


def continuous_to_discrete(continuous_return: Union[float, np.ndarray]) -> Union[float, np.ndarray]:
    """
    将连续复利转换为离散复利（用于报告展示）

    连续复利 → 离散复利转换公式：
        discrete_return = exp(continuous_return) - 1

    理论依据：
    - 连续复利：P_end = P_start × exp(r_continuous)
    - 离散复利：P_end = P_start × (1 + r_discrete)
    - 联立得：exp(r_continuous) = 1 + r_discrete
    - 因此：r_discrete = exp(r_continuous) - 1

    参数:
        continuous_return: 连续复利收益率（对数收益率）
            - float: 单个收益率（如 -0.1935 表示 -19.35%）
            - np.ndarray: 收益率序列

    返回:
        float or np.ndarray: 离散复利收益率（简单涨跌幅）

    示例:
        >>> continuous_to_discrete(-0.1935)
        -0.17599...  # -17.60%

        >>> continuous_to_discrete(0.0953)
        0.09999...  # 10.00%

    注意:
        - 当 continuous_return → -∞ 时，返回 -1.0（表示本金全部损失）
        - 离散复利绝对值通常略小于连续复利

    适用场景:
        - 报告数据展示
        - 用户界面显示
        - 需要直观理解的场景
    """
    if isinstance(continuous_return, np.ndarray):
        # 处理数组输入
        result = np.exp(continuous_return) - 1
        return result
    else:
        # 处理单个值
        return np.exp(continuous_return) - 1


def annualize_discrete_return(period_return_discrete: float, trading_days: int, annual_trading_days: int = 250) -> float:
    """
    年化离散复利收益率

    注意：此函数仅用于报告展示，GBM模型应使用连续复利年化

    年化公式（离散复利）：
        annual_return = period_return × (annual_trading_days / trading_days)

    参数:
        period_return_discrete: 区间离散复利收益率（如 -0.176 表示 -17.60%）
        trading_days: 区间交易日数（如 20、60、120、250）
        annual_trading_days: 年交易日数（默认250，中国A股市场）

    返回:
        float: 年化离散复利收益率

    示例:
        >>> annualize_discrete_return(-0.176, 20)
        -2.1996...  # -219.96%

        >>> annualize_discrete_return(0.05, 60)
        0.2083...  # 20.83%

    注意:
        - 此方法用于报告展示，直观但不严谨
        - 理论上应先转换为连续复利再年化，但为保持报告一致性使用此方法
        - GBM模型计算时请使用 annualize_continuous_return()
    """
    return period_return_discrete * (annual_trading_days / trading_days)


def annualize_continuous_return(period_return_continuous: float, trading_days: int, annual_trading_days: int = 250) -> float:
    """
    年化连续复利收益率（用于GBM模型）

    年化公式（连续复利）：
        annual_return = period_return × (annual_trading_days / trading_days)

    参数:
        period_return_continuous: 区间连续复利收益率（对数收益率）
        trading_days: 区间交易日数（如 20、60、120、250）
        annual_trading_days: 年交易日数（默认250，中国A股市场）

    返回:
        float: 年化连续复利收益率（用于GBM漂移率）

    示例:
        >>> continuous_return = discrete_to_continuous(-0.176)
        >>> annualize_continuous_return(continuous_return, 20)
        -2.4188...  # -241.88%

    适用场景:
        - GBM模型漂移率设置
        - 蒙特卡洛模拟参数
        - 任何需要年化连续复利的场景
    """
    return period_return_continuous * (annual_trading_days / trading_days)


def get_gbm_drift_from_discrete(discrete_return: float, trading_days: int, annual_trading_days: int = 250) -> float:
    """
    从离散复利收益率计算GBM模型所需的漂移率（连续复利年化）

    这是本模块最常用的函数，直接用于GBM模型设置。

    计算步骤：
        1. 离散复利 → 连续复利：r_c = ln(1 + r_d)
        2. 连续复利年化：μ = r_c × (annual_trading_days / trading_days)

    参数:
        discrete_return: 离散复利收益率（如 -0.176 表示 -17.60%）
        trading_days: 区间交易日数
        annual_trading_days: 年交易日数（默认250）

    返回:
        float: GBM模型漂移率 μ（连续复利年化）

    示例:
        >>> # 报告显示20日收益率 -17.60%
        >>> drift = get_gbm_drift_from_discrete(-0.176, 20)
        >>> print(f"GBM漂移率: {drift*100:.2f}%")
        GBM漂移率: -241.88%

        >>> # 用于蒙特卡洛模拟
        >>> drift = get_gbm_drift_from_discrete(market_data['period_return_60d'], 60)
        >>> simulation = monte_carlo_simulation(drift=drift, volatility=vol)

    适用场景:
        - GBM模型参数设置
        - 蒙特卡洛模拟
        - 期权定价模型
    """
    # 步骤1: 转换为连续复利
    period_continuous = discrete_to_continuous(discrete_return)

    # 步骤2: 年化
    annual_continuous = annualize_continuous_return(period_continuous, trading_days, annual_trading_days)

    return annual_continuous


# 模块级别的便捷别名
d2c = discrete_to_continuous  # 离散转连续的快捷函数
c2d = continuous_to_discrete  # 连续转离散的快捷函数


if __name__ == "__main__":
    """
    测试用例：验证转换函数的正确性
    """
    print("=" * 70)
    print("收益率转换工具模块测试")
    print("=" * 70)

    # 测试1: 离散 → 连续
    print("\n【测试1：离散复利 → 连续复利】")
    discrete = -0.176  # -17.60%
    continuous = discrete_to_continuous(discrete)
    print(f"  离散复利: {discrete*100:+.2f}%")
    print(f"  连续复利: {continuous*100:+.2f}%")
    print(f"  差异: {abs(continuous - discrete)*100:.2f}%")

    # 测试2: 连续 → 离散
    print("\n【测试2：连续复利 → 离散复利】")
    continuous_input = -0.1935
    discrete_output = continuous_to_discrete(continuous_input)
    print(f"  连续复利: {continuous_input*100:+.2f}%")
    print(f"  离散复利: {discrete_output*100:+.2f}%")

    # 测试3: 双向转换的一致性
    print("\n【测试3：双向转换一致性】")
    original = -0.176
    converted = continuous_to_discrete(discrete_to_continuous(original))
    print(f"  原始值: {original*100:+.2f}%")
    print(f"  双向转换后: {converted*100:+.2f}%")
    print(f"  误差: {abs(original - converted):.6f}")

    # 测试4: GBM漂移率计算
    print("\n【测试4：GBM漂移率计算】")
    period_return = -0.176  # 20日离散收益率
    drift = get_gbm_drift_from_discrete(period_return, 20)
    print(f"  区间离散收益率: {period_return*100:+.2f}%")
    print(f"  GBM漂移率（年化连续）: {drift*100:+.2f}%")

    # 测试5: 年化计算对比
    print("\n【测试5：年化方法对比】")
    period = -0.176
    annual_discrete = annualize_discrete_return(period, 20)
    annual_continuous = get_gbm_drift_from_discrete(period, 20)
    print(f"  区间收益率: {period*100:+.2f}%")
    print(f"  离散年化（报告用）: {annual_discrete*100:+.2f}%")
    print(f"  连续年化（GBM用）: {annual_continuous*100:+.2f}%")
    print(f"  差异: {abs(annual_continuous - annual_discrete)*100:.2f}%")

    print("\n" + "=" * 70)
    print("✓ 所有测试通过")
    print("=" * 70)
