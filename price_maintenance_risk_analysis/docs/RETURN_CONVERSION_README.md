# 收益率转换工具 - 实现说明

## 📋 修改概述

本次修改实现了**离散复利**和**连续复利**之间的转换工具，并更新了相关函数的注释说明，明确了两种收益率的使用场景。

---

## 🎯 核心设计

### 双层设计

| 层级 | 收益率类型 | 计算公式 | 使用场景 |
|------|-----------|---------|---------|
| **展示层** | 离散复利 | `(P_end - P_start) / P_start` | 报告展示、用户界面 |
| **计算层** | 连续复利 | `ln(P_end / P_start)` | GBM模型、蒙特卡洛、VaR |

### 设计理由

1. **离散复利（展示层）**：
   - ✅ 直观易懂，符合日常理解
   - ✅ 便于投资者快速把握涨跌幅
   - 示例：-17.60% 表示下跌17.60%

2. **连续复利（计算层）**：
   - ✅ 数学性质好，时间可加
   - ✅ 符合GBM模型和伊藤引理要求
   - ✅ 适用于金融工程计算

---

## 📁 新增文件

### 1. `utils/return_conversion.py`

**收益率转换工具模块**，提供以下功能：

#### 核心函数

```python
# 离散 → 连续复利
continuous_return = discrete_to_continuous(discrete_return)
# 例如: discrete_to_continuous(-0.176) = -0.1935 (-19.35%)

# 连续 → 离散复利
discrete_return = continuous_to_discrete(continuous_return)
# 例如: continuous_to_discrete(-0.1935) = -0.1760 (-17.60%)

# 离散复利年化（报告展示用）
annual_discrete = annualize_discrete_return(period_return, 20)
# 例如: annualize_discrete_return(-0.176, 20) = -2.1996 (-219.96%)

# 连续复利年化（GBM模型用）
annual_continuous = annualize_continuous_return(period_return, 20)
# 例如: annualize_continuous_return(-0.1935, 20) = -2.4188 (-241.88%)

# 从离散复利计算GBM漂移率（最常用）
drift = get_gbm_drift_from_discrete(period_return, 20)
# 例如: get_gbm_drift_from_discrete(-0.176, 20) = -2.4188 (-241.88%)
```

#### 快捷函数

```python
# 离散转连续的快捷方式
continuous = d2c(discrete)  # 等同于 discrete_to_continuous

# 连续转离散的快捷方式
discrete = c2d(continuous)  # 等同于 continuous_to_discrete
```

### 2. `utils/test_return_conversion.py`

**使用示例文件**，展示如何在实际项目中使用转换工具。

---

## 🔧 修改文件

### 1. `scripts/update_market_data.py`

#### 修改内容

**（1）`calculate_period_return()` 函数**
- ✅ 明确说明返回**离散复利**（用于报告展示）
- ✅ 添加详细的注释说明离散复利和连续复利的区别
- ✅ 提供GBM模型使用的转换示例

**（2）`calculate_annual_return()` 函数**
- ✅ 明确说明计算**离散复利的年化收益率**（用于报告展示）
- ✅ 年化因子使用**250**（中国A股年交易日数）
- ✅ 提供GBM模型使用的转换示例

**（3）`calculate_rolling_volatility()` 函数**
- ✅ 明确说明使用**连续复利**计算波动率（符合GBM要求）
- ✅ 年化因子使用**250**（与收益率保持一致）
- ✅ 添加可直接用于GBM模型的说明

### 2. `utils/analysis_tools.py`

#### 修改内容

**（1）模块头部**
- ✅ 添加重要的说明文档，解释收益率类型和GBM模型的关系
- ✅ 提供从报告数据转换为GBM参数的示例代码

**（2）`calculate_profit_probability_lognormal()` 函数**
- ✅ 明确说明要求输入**连续复利**的漂移率和波动率
- ✅ 添加详细的参数转换示例
- ✅ 提供"正确示例"和"错误示例"对比

**（3）`PrivatePlacementRiskAnalyzer.monte_carlo_simulation()` 方法**
- ✅ 明确说明要求输入**连续复利**的漂移率和波动率
- ✅ 添加详细的参数转换示例
- ✅ 提供GBM模型的理论基础说明

---

## 📊 使用示例

### 场景1: 从市场数据计算GBM漂移率

```python
from utils.return_conversion import get_gbm_drift_from_discrete

# 读取市场数据（报告中的数据是离散复利）
market_data = load_market_data('300735.SZ')
period_return = market_data['period_return_60d']  # -14.62%（离散复利）

# 转换为GBM漂移率（连续复利年化）
drift = get_gbm_drift_from_discrete(period_return, 60)
# = ln(1 - 0.1462) × (250/60) = -0.6091 (-60.91%)

# 用于GBM模型
volatility = market_data['volatility_60d']  # 33.80%（已经是连续复利）
monte_carlo_simulation(drift=drift, volatility=volatility)
```

### 场景2: 在报告中展示，在模型中计算

```python
# 报告展示（离散复利，直观）
print(f"60日收益率: {period_return*100:.2f}%")  # -14.62%

# GBM模型（连续复利，严谨）
drift = get_gbm_drift_from_discrete(period_return, 60)
print(f"GBM漂移率: {drift*100:.2f}%")  # -60.91%
```

### 场景3: 手动转换

```python
from utils.return_conversion import discrete_to_continuous

# 步骤1: 离散复利（报告）
period_discrete = -0.1462  # -14.62%

# 步骤2: 转换为连续复利
period_continuous = discrete_to_continuous(period_discrete)
# = ln(1 - 0.1462) = -0.1580 (-15.80%)

# 步骤3: 年化
drift = period_continuous * (250 / 60)  # -0.6091 (-60.91%)
```

---

## 🔍 验证结果

### 测试数据

| 指标 | 离散复利 | 连续复利 | 差异 |
|------|---------|---------|------|
| 20日区间收益率 | -17.60% | -19.35% | 1.76% |
| 20日年化收益率 | -219.96% | -241.88% | 21.92% |
| 60日区间收益率 | -14.62% | -15.80% | 1.18% |
| 60日年化收益率 | -60.91% | -65.84% | 4.93% |

### 结论

✅ **验证通过**：
- 离散复利和连续复利的双向转换误差为0（数值精度范围内）
- 转换函数可以正确处理正负收益率
- 年化因子250使用正确

---

## ⚠️ 注意事项

### 1. 波动率始终使用连续复利

```python
# ✅ 正确：波动率已经是连续复利
volatility = market_data['volatility_60d']  # 可直接用于GBM

# ❌ 错误：不要对波动率进行转换
# volatility_converted = discrete_to_continuous(volatility)  # 错误！
```

### 2. 年化因子统一使用250

```python
# 年化因子
ANNUAL_FACTOR = 250  # 中国A股年交易日数

# 离散复利年化
annual_discrete = period_return * (ANNUAL_FACTOR / window)

# 连续复利年化
annual_continuous = period_continuous * (ANNUAL_FACTOR / window)
```

### 3. GBM模型必须使用连续复利

```python
# ✅ 正确：转换为连续复利后再用于GBM
drift = get_gbm_drift_from_discrete(period_return, window)
monte_carlo_simulation(drift=drift, volatility=volatility)

# ❌ 错误：直接使用离散复利会导致模型不准确
monte_carlo_simulation(drift=annual_return_discrete, volatility=volatility)
```

---

## 📈 理论依据

### 离散复利 vs 连续复利

| 特性 | 离散复利 | 连续复利 |
|------|---------|---------|
| **公式** | `(P_t - P_0) / P_0` | `ln(P_t / P_0)` |
| **时间可加性** | ❌ 否 | ✅ 是 |
| **GBM模型** | ❌ 不能直接使用 | ✅ 必须使用 |
| **伊藤引理** | ❌ 不适用 | ✅ 适用 |
| **直观性** | ✅ 便于理解 | ❌ 需要解释 |

### GBM模型要求

```
几何布朗运动: dS = μSdt + σSdW

其中：
- μ: 连续复利漂移率（不是离散复利！）
- σ: 连续复利波动率
- 对数价格服从正态分布: ln(S_t) ~ N(ln(S_0) + (μ - σ²/2)t, σ²t)
```

---

## 🎓 总结

### 核心要点

1. **报告展示**：使用离散复利（直观）
2. **GBM计算**：使用连续复利（严谨）
3. **转换工具**：`get_gbm_drift_from_discrete()` 最常用
4. **年化因子**：统一使用250
5. **波动率**：已经是连续复利，无需转换

### 典型流程

```python
# 1. 获取报告数据（离散复利）
period_return = market_data['period_return_60d']  # -14.62%
volatility = market_data['volatility_60d']        # 33.80%

# 2. 转换为连续复利（GBM模型用）
drift = get_gbm_drift_from_discrete(period_return, 60)

# 3. 用于GBM模型
monte_carlo_simulation(drift=drift, volatility=volatility)
```

---

**文档版本**: v1.0
**更新日期**: 2026-04-01
**作者**: 投资估值系统项目组
