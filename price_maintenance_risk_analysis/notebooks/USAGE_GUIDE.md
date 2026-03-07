# 定增分析 - 使用真实数据指南

## 概述

所有分析脚本现在都支持使用真实市场数据（波动率、收益率等），而不是假设的参数。

## 核心模块

### 1. utils/config_loader.py
**统一的配置加载模块**，自动加载：
- ✅ 定增参数（从JSON文件）
- ✅ 市场数据（从JSON文件，如果存在）
- ✅ 优先使用真实数据，回退到参数文件

## 使用方法

### 方式1: 在Notebook中使用（推荐）

```python
# 导入配置加载器
from utils.config_loader import load_placement_config, print_config_summary
from utils.analysis_tools import PrivatePlacementRiskAnalyzer

# 一行代码加载配置
project_params, risk_params = load_placement_config('300735.SZ')

# 查看配置摘要
print_config_summary(project_params, risk_params)

# 创建分析器
analyzer = PrivatePlacementRiskAnalyzer(**project_params)

# 使用真实参数运行分析
sim_results = analyzer.monte_carlo_simulation(
    n_simulations=10000,
    volatility=risk_params['volatility'],  # 真实波动率
    drift=risk_params['drift'],             # 真实收益率
    seed=42
)
```

### 方式2: 在Python脚本中使用

```python
#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
sys.path.append('..')

from utils.config_loader import load_placement_config
from utils.analysis_tools import PrivatePlacementRiskAnalyzer

def analyze_private_placement(stock_code='300735.SZ'):
    """分析定增项目"""

    # 1. 加载配置（自动使用真实数据）
    project_params, risk_params = load_placement_config(stock_code)

    # 2. 创建分析器
    analyzer = PrivatePlacementRiskAnalyzer(**project_params)

    # 3. 运行分析
    # ... 你的分析逻辑 ...

    print(f"波动率: {risk_params['volatility']*100:.2f}%")
    print(f"收益率: {risk_params['drift']*100:.2f}%")

if __name__ == '__main__':
    analyze_private_placement('300735.SZ')
```

### 方式3: 快速加载

```python
from utils.config_loader import quick_load

# 一行代码加载并打印摘要
project_params, risk_params = quick_load('300735.SZ')
```

## 已更新的Notebook

以下notebook已更新支持真实数据：

| Notebook | 使用真实数据 | 数据来源 |
|----------|-------------|---------|
| 01_sensitivity_analysis.ipynb | ✅ | 自动加载 |
| 02_stress_test.ipynb | ✅ | 自动加载 |
| 04_monte_carlo_simulation.ipynb | ✅ | 自动加载 |
| 05_var_calculation.ipynb | ✅ | 自动加载 |
| 07_market_data_analysis.ipynb | ✅ | 生成数据 |

## 数据文件说明

### 300735_SZ_placement_params.json
定增参数文件（由fetch_gh_data.py生成）：
```json
{
  "issue_price": 20.76,
  "current_price": 25.95,
  "lockup_period": 6,
  "volatility": 0.3117,
  ...
}
```

### 300735_SZ_market_data.json
市场数据文件（由07_market_data_analysis.ipynb生成）：
```json
{
  "volatility_30d": 0.3512,
  "volatility_60d": 0.3117,
  "volatility_120d": 0.2856,
  "volatility_180d": 0.2734,
  "volatility": 0.3117,
  "drift": 0.0845,
  "annual_return_60d": 0.0845,
  ...
}
```

## 工作流程

### 首次分析新股票

```bash
# 1. 获取定增参数
cd price_maintenance_risk_analysis
python fetch_gh_data.py

# 2. 生成市场数据（可选但推荐）
jupyter notebook notebooks/07_market_data_analysis.ipynb
# 运行所有cell，生成 300735_SZ_market_data.json

# 3. 运行任何分析notebook
jupyter notebook notebooks/04_monte_carlo_simulation.ipynb
# 会自动使用真实市场数据
```

### 切换分析标的

```python
# 只需修改股票代码
project_params, risk_params = load_placement_config('000001.SZ')  # 平安银行
# 或
project_params, risk_params = load_placement_config('600519.SH')  # 贵州茅台
```

## 数据优先级

```
市场数据（如果存在）→ 定增参数文件 → 默认假设
    ↓                 ↓              ↓
  真实计算           参数文件        vol=0.35
  (推荐)            (备选)          drift=0.08
```

## 返回的参数说明

### project_params（定增项目参数）
```python
{
    'issue_price': 20.76,        # 发行价格（元/股）
    'issue_shares': 5000000,     # 发行数量（股）
    'lockup_period': 6,          # 锁定期（月）
    'current_price': 25.95,      # 当前价格（元/股）
    'risk_free_rate': 0.03       # 无风险利率
}
```

### risk_params（风险参数）
```python
{
    'volatility': 0.3117,         # 默认波动率（60日）
    'drift': 0.0845,              # 默认收益率（60日年化）
    'volatility_30d': 0.3512,     # 30日波动率（如果有市场数据）
    'volatility_60d': 0.3117,     # 60日波动率
    'volatility_120d': 0.2856,    # 120日波动率
    'volatility_180d': 0.2734,    # 180日波动率
    'data_source': 'market_data'   # 数据来源标识
}
```

## 常见使用场景

### 场景1: 蒙特卡洛模拟

```python
from utils.config_loader import load_placement_config
from utils.analysis_tools import PrivatePlacementRiskAnalyzer

project_params, risk_params = load_placement_config('300735.SZ')
analyzer = PrivatePlacementRiskAnalyzer(**project_params)

sim_results = analyzer.monte_carlo_simulation(
    n_simulations=10000,
    volatility=risk_params['volatility'],  # 使用真实波动率
    drift=risk_params['drift'],             # 使用真实收益率
    seed=42
)
```

### 场景2: VaR计算

```python
from utils.config_loader import load_placement_config

project_params, risk_params = load_placement_config('300735.SZ')

# 使用真实参数计算VaR
volatility = risk_params['volatility']
drift = risk_params['drift']

# VaR计算逻辑...
var_95 = drift - 1.645 * volatility / np.sqrt(252)
```

### 场景3: 敏感性分析

```python
from utils.config_loader import load_placement_config

project_params, risk_params = load_placement_config('300735.SZ')

# 使用不同窗口的波动率
for window in ['30d', '60d', '120d']:
    vol = risk_params[f'volatility_{window}']
    print(f"{window} 波动率: {vol*100:.2f}%")
```

## 完整示例：自定义分析脚本

```python
#!/usr/bin/env python
# custom_analysis.py
import sys
sys.path.append('..')

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from utils.config_loader import load_placement_config, print_config_summary
from utils.analysis_tools import PrivatePlacementRiskAnalyzer

def custom_analysis(stock_code='300735.SZ'):
    """自定义分析示例"""

    # 1. 加载配置
    print("加载配置...")
    project_params, risk_params = load_placement_config(stock_code)
    print_config_summary(project_params, risk_params)

    # 2. 创建分析器
    analyzer = PrivatePlacementRiskAnalyzer(**project_params)

    # 3. 你的自定义分析
    print(f"\n使用真实参数进行分析:")
    print(f"  波动率: {risk_params['volatility']*100:.2f}%")
    print(f"  收益率: {risk_params['drift']*100:.2f}%")

    # 示例：计算盈亏平衡价格
    target_return = 0.20
    break_even_price = analyzer.calculate_break_even_price(target_return)

    print(f"\n{target_return*100:.0f}%年化收益率盈亏平衡价: {break_even_price:.2f} 元")
    print(f"当前价格: {project_params['current_price']:.2f} 元")

    # 4. 使用多窗口波动率进行对比分析
    if 'volatility_30d' in risk_params:
        print(f"\n不同窗口波动率对比:")
        for window in ['30d', '60d', '120d', '180d']:
            vol = risk_params.get(f'volatility_{window}')
            if vol:
                print(f"  {window}: {vol*100:.2f}%")

if __name__ == '__main__':
    custom_analysis('300735.SZ')
```

## 故障排除

### Q1: 加载失败，提示文件不存在

```python
# 确保先生成数据文件
# 1. 运行 fetch_gh_data.py
python fetch_gh_data.py

# 2. 运行 07_market_data_analysis.ipynb
jupyter notebook notebooks/07_market_data_analysis.ipynb
```

### Q2: 想使用不同的波动率窗口

```python
project_params, risk_params = load_placement_config('300735.SZ')

# 使用30日波动率（更敏感）
volatility_30d = risk_params['volatility_30d']

# 使用180日波动率（更稳定）
volatility_180d = risk_params['volatility_180d']
```

### Q3: 想分析其他股票

```python
# 只需修改股票代码
project_params, risk_params = load_placement_config('000001.SZ')  # 平安银行

# 确保该股票的数据文件存在：
# - 000001_SH_placement_params.json
# - 000001_SH_market_data.json（可选）
```

## 总结

现在所有分析脚本都可以：
1. ✅ **一行代码加载配置** - `load_placement_config('300735.SZ')`
2. ✅ **自动使用真实数据** - 波动率、收益率来自历史计算
3. ✅ **智能回退机制** - 数据不存在时使用参数文件
4. ✅ **统一数据格式** - 所有脚本使用相同的参数结构

不再需要手动假设 `volatility=0.35`、`drift=0.08` 等参数！
