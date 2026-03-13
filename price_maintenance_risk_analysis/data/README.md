# 数据目录说明

本目录是`price_maintenance_risk_analysis`模块的唯一数据目录。

## 目录结构

```
price_maintenance_risk_analysis/
├── data/                              # 统一数据目录（唯一）
│   ├── 300735_SZ_placement_params.json   # 定增参数配置
│   ├── 300735_SZ_market_data.json        # 市场数据（波动率、收益率等）
│   ├── 300735_SZ_industry_data.json      # 行业指数数据
│   ├── market_indices_scenario_data.json # 市场指数情景数据
│   └── market_indices_scenario_data_v2.json # 市场指数情景数据v2
├── scripts/                           # 脚本目录
│   ├── generate_word_report_v2.py     # 报告生成脚本
│   └── update_market_data.py          # 数据更新脚本
└── utils/                             # 工具类目录
    ├── config_loader.py               # 配置加载器
    └── market_data_loader.py          # 市场数据加载器
```

## 数据文件说明

### 1. 定增参数配置
- **文件**: `300735_SZ_placement_params.json`
- **内容**: 定增项目的基本参数（发行价、锁定期、融资金额等）
- **生成**: 手动创建或由脚本自动生成

### 2. 市场数据
- **文件**: `300735_SZ_market_data.json`
- **内容**: 个股市场数据（波动率、收益率、移动平均线等）
- **生成**: `update_market_data.py`
- **更新频率**: 建议每日更新

### 3. 行业指数数据
- **文件**: `300735_SZ_industry_data.json`
- **内容**: 申万行业指数数据（行业分类、波动率、收益率等）
- **生成**: `update_market_data.py`
- **更新频率**: 建议每日更新

### 4. 市场指数情景数据
- **文件**: `market_indices_scenario_data.json`
- **内容**: 市场指数在不同情景下的数据
- **生成**: `update_indices_data.py`

## 路径引用规范

### 在Python脚本中引用数据目录

**方法1：使用相对路径（推荐）**
```python
# 在scripts/目录下的脚本
import os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
DATA_DIR = os.path.join(PROJECT_DIR, 'data')

# 使用DATA_DIR
data_file = os.path.join(DATA_DIR, '300735_SZ_market_data.json')
```

**方法2：使用工具函数（最简单）**
```python
# 自动检测数据目录
from utils.config_loader import load_placement_config
from utils.market_data_loader import load_market_data

# data_dir参数默认为'data'，会自动查找
project_params, risk_params, market_data = load_placement_config('300735.SZ')
market_data = load_market_data('300735.SZ')  # 默认从data目录加载
```

**方法3：使用绝对路径**
```python
import os
# 获取脚本所在目录的绝对路径
script_dir = os.path.dirname(os.path.abspath(__file__))
# data目录在项目根目录下
data_dir = os.path.join(os.path.dirname(script_dir), 'data')
```

## 数据更新流程

### 自动更新市场数据
```bash
cd price_maintenance_risk_analysis/scripts
python update_market_data.py
```

这会自动更新：
- 个股市场数据 (`300735_SZ_market_data.json`)
- 行业指数数据 (`300735_SZ_industry_data.json`)
- 定增参数配置 (`300735_SZ_placement_params.json`，如果不存在则创建)

### 手动配置定增参数
编辑 `300735_SZ_placement_params.json`：
```json
{
  "stock_code": "300735.SZ",
  "stock_name": "光弘科技",
  "issue_price": 22.21,
  "current_price": 22.88,
  "lockup_period": 6,
  "financing_amount": 7.0
}
```

## 注意事项

1. **不要使用** `../data` 或 `../../data` 等相对路径
2. **推荐使用** 工具函数 `load_market_data()` 和 `load_placement_config()`
3. **数据目录统一**: 所有数据文件都在 `price_maintenance_risk_analysis/data/`
4. **路径优先级**: 工具函数会自动在以下位置查找数据：
   - `data/` (price_maintenance_risk_analysis/data) - 最高优先级
   - `.` (当前目录)
   - `..` (父目录，兼容性)

## 清理说明

**已删除**: 项目根目录的 `/data/` 目录（重复，已废弃）

**原因**:
- 与 `price_maintenance_risk_analysis/data/` 重复
- 容易导致脚本路径混乱
- 统一数据管理，避免歧义

**迁移说明**:
- 所有脚本已更新为使用 `price_maintenance_risk_analysis/data/`
- 工具函数已优化路径检测逻辑
- 无需手动修改任何配置
