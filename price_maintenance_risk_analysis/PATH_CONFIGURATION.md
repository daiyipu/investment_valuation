# 数据路径配置说明

## ✅ 已统一的数据目录结构

**唯一数据目录**: `price_maintenance_risk_analysis/data/`

所有脚本和工具函数现在都正确使用这个目录。

## 路径自动检测机制

### 工具函数的智能路径检测

**utils/config_loader.py** 和 **utils/market_data_loader.py** 会按以下优先级自动检测数据目录：

1. `'data/'` - 当从项目根目录运行时
2. `'../data/'` - 当从 scripts/ 目录运行时
3. `'./'` - 当前目录（兼容性）
4. `'../'` - 父目录（兼容性）

### 实际使用情况

| 运行目录 | 实际使用路径 | 状态 |
|---------|-------------|------|
| 从 `scripts/` 运行 | `../data/` | ✅ 正常 |
| 从 `price_maintenance_risk_analysis/` 运行 | `data/` | ✅ 正常 |
| 使用绝对路径运行 | 自动检测 `../data/` | ✅ 正常 |

## 验证结果

### 测试1: 从 scripts/ 目录
```bash
cd price_maintenance_risk_analysis/scripts
python -c "from utils.config_loader import load_placement_config; ..."
```
**结果**: ✅ 使用 `../data/300735_SZ_placement_params.json`

### 测试2: 从项目根目录
```bash
cd price_maintenance_risk_analysis
python -c "from utils.config_loader import load_placement_config; ..."
```
**结果**: ✅ 使用 `data/300735_SZ_placement_params.json`

### 测试3: 实际脚本运行
```bash
python generate_word_report_v2.py
```
**结果**: ✅ 成功生成报告

## 当前数据文件清单

```
price_maintenance_risk_analysis/data/
├── 300735_SZ_placement_params.json        (4.3KB) ✅
├── 300735_SZ_market_data.json             (826B)  ✅
├── 300735_SZ_industry_data.json           (953B)  ✅
├── market_indices_scenario_data.json       (6.4KB) ✅
└── market_indices_scenario_data_v2.json    (3.8KB) ✅
```

## 关于PE历史数据获取失败的说明

您看到的警告信息：
```
⚠️ 未获取到消费电子零部件及组装(850854.SI)的历史PE数据
```

**这不是路径问题**，而是数据源问题：

1. **原因**: Tushare API 的申万行业指数历史PE数据不可用
2. **影响**: 无法生成行业PE历史分位数趋势图
3. **处理**: 脚本已正确处理，跳过该图表，继续生成报告
4. **建议**: 个股PE数据仍可正常使用，不影响其他分析

## 已删除的重复目录

**已删除**: `/data/` (项目根目录)
- 原因: 与 `price_maintenance_risk_analysis/data/` 重复
- 状态: ✅ 已删除，无影响

## 路径配置最佳实践

### ✅ 推荐做法

**方法1: 使用工具函数（最简单）**
```python
from utils.config_loader import load_placement_config
from utils.market_data_loader import load_market_data

# data_dir参数默认为None，自动检测
project_params, risk_params, market_data = load_placement_config('300735.SZ')
data = load_market_data('300735.SZ')  # 自动查找
```

**方法2: 使用PROJECT_DIR常量**
```python
import os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
DATA_DIR = os.path.join(PROJECT_DIR, 'data')

data_file = os.path.join(DATA_DIR, 'file.json')
```

### ❌ 不推荐做法

```python
# ❌ 硬编码相对路径（容易出错）
data_file = '../data/file.json'

# ❌ 使用固定的data_dir='data'
def my_func(data_dir='data'):  # 不会自动检测
    ...
```

## 总结

✅ **所有脚本已更新为统一的数据目录配置**
✅ **工具函数支持从任何目录自动检测数据路径**
✅ **路径配置已全面测试验证通过**
✅ **PE历史数据获取失败是API限制，不是路径问题**

**数据目录统一**: `price_maintenance_risk_analysis/data/`
