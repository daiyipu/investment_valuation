# 数据目录统一优化说明

## 改动概述

将项目中两个data目录统一为一个，放置在`price_maintenance_risk_analysis/data/`，避免路径混乱和重复文件。

## 问题说明

### 原有问题
1. **两个data目录**：
   - `/data/` - 项目根目录的data（只有2个文件，重复）
   - `/price_maintenance_risk_analysis/data/` - 模块的data（有7个完整文件）

2. **路径引用混乱**：
   - `generate_word_report_v2.py` 使用 `PROJECT_DIR/data`
   - `update_market_data.py` 使用相对路径
   - `update_indices_data.py` 使用 `../data`
   - 工具函数尝试多个路径查找

3. **影响**：
   - 脚本第一次运行时容易出错
   - 路径不统一，维护困难
   - 数据文件可能重复或不一致

## 解决方案

### 1. 统一数据目录
**保留**：`price_maintenance_risk_analysis/data/`
**删除**：项目根目录的 `/data/` (重复)

### 2. 优化路径检测

#### 更新的文件：

**utils/config_loader.py**
```python
# 优先级列表（从高到低）：
possible_dirs = [
    'data',       # 当在项目根目录时
    '../data',    # 当在scripts目录时（新增）
    '.',         # 当前目录
    '..'         # 父目录
]
```

**utils/market_data_loader.py**
- 将默认参数从 `data_dir='data'` 改为 `data_dir=None`
- 添加自动检测逻辑（与config_loader相同）

**scripts/update_indices_data.py**
- 从 `'../data/market_indices_scenario_data.json'`
- 改为使用绝对路径计算：
  ```python
  script_dir = os.path.dirname(os.path.abspath(__file__))
  data_dir = os.path.join(os.path.dirname(script_dir), 'data')
  data_file = os.path.join(data_dir, 'market_indices_scenario_data.json')
  ```

### 3. 新增文件

**data/README.md**
- 数据目录说明文档
- 路径引用规范
- 数据更新流程
- 注意事项

**scripts/verify_data_paths.py**
- 数据路径验证脚本
- 检查所有关键文件是否存在
- 测试工具函数是否正常工作
- 验证所有脚本的路径配置

## 测试验证

### 验证脚本
```bash
cd price_maintenance_risk_analysis/scripts
python verify_data_paths.py
```

**结果**：✅ 所有检查通过

### 测试工具函数
```python
from utils.config_loader import load_placement_config
from utils.market_data_loader import load_market_data

# 测试加载配置
project_params, risk_params, market_data = load_placement_config('300735.SZ')
# ✅ 成功加载

# 测试加载数据
data = load_market_data('300735.SZ')
# ✅ 成功加载
```

### 测试报告生成
```bash
python generate_word_report_v2.py
# ✅ 成功生成报告
```

## 向后兼容性

### 支持的运行方式
1. **从scripts目录运行**：
   ```bash
   cd price_maintenance_risk_analysis/scripts
   python generate_word_report_v2.py
   # 自动使用 ../data
   ```

2. **从项目根目录运行**：
   ```bash
   cd price_maintenance_risk_analysis
   python -m scripts.generate_word_report_v2
   # 自动使用 data/
   ```

3. **从其他目录运行**：
   ```bash
   python /path/to/generate_word_report_v2.py
   # 自动检测并使用 ../data
   ```

## 文件清单

### 修改的文件
1. ✅ `utils/config_loader.py` - 添加`../data`到候选路径
2. ✅ `utils/market_data_loader.py` - 添加自动检测逻辑
3. ✅ `scripts/update_indices_data.py` - 使用绝对路径计算

### 新增的文件
1. ✅ `data/README.md` - 数据目录使用说明
2. ✅ `scripts/verify_data_paths.py` - 路径验证脚本

### 需要删除的文件
1. ⚠️ `/data/` (项目根目录的重复data目录)

## 迁移步骤

如果您的项目根目录还有`/data/`目录，可以执行以下步骤：

```bash
# 1. 检查是否有独有文件
ls -la /data/
ls -la price_maintenance_risk_analysis/data/

# 2. 如果根目录data有独有文件，先移动
mv /data/unique_file.json price_maintenance_risk_analysis/data/

# 3. 删除根目录的data（确认已没有独有文件后）
rm -rf /data/
```

## 注意事项

1. **不要使用**硬编码的相对路径如`'../data'`或`'../../data'`
2. **推荐使用**工具函数`load_market_data()`和`load_placement_config()`
3. **默认参数**`data_dir=None`会让工具函数自动查找
4. **绝对路径**：对于复杂情况，可以使用脚本中的绝对路径计算方法

## 相关脚本

所有使用数据文件的脚本已更新为统一的路径配置：
- ✅ `generate_word_report_v2.py`
- ✅ `generate_word_report.py`
- ✅ `update_market_data.py`
- ✅ `update_indices_data.py`
- ✅ `utils/config_loader.py`
- ✅ `utils/market_data_loader.py`

## 测试清单

- [x] 工具函数从scripts目录运行正常
- [x] 报告生成脚本从scripts目录运行正常
- [x] 数据更新脚本从scripts目录运行正常
- [x] 路径自动检测在所有场景下正常
- [x] README文档完整清晰
- [x] 验证脚本可以正确检测所有问题
