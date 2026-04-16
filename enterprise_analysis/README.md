# 企业财务信息分析评价系统1.3.32版本

## 项目概述

这是一个企业财务信息分析评价系统，支持导入企业财务报表数据，计算财务指标，生成财务评分，并进行趋势分析和历史数据查询。

## 🆕 1.3.32版本新增功能

### 用户认证和数据隔离
- ✅ 多用户登录认证系统（用户名+密码）
- ✅ 数据按用户完全隔离（企业私有数据）
- ✅ 行业阈值全局共享（所有用户受益）
- ✅ 密码SHA256哈希加密存储

### 功能增强
- ✅ 手动填写财务数据模式
- ✅ 用户管理工具（创建、删除、重置密码）
- ✅ 步骤跳转优化（一次点击即可跳转）
- ✅ 访问地址优化（显示localhost而非0.0.0.0）

### 文档完善
- ✅ 完整的安装部署指南
- ✅ 统信UOS ARM64部署指南
- ✅ 用户数据隔离设计说明
- ✅ 统一的依赖管理（requirements.txt）

## 🆕 1.3.27版本新增功能

- **Tushare上市公司模式**: 新增直接从Tushare API获取上市公司官方财务数据的功能
- **财务指标数据源优化**: 修复财务指标数据源问题，优先使用Tushare直接获取的指标数据
  - 修复 inv_turn（存货周转率）错误：0 → 3.1467 ✅
  - 修复 grossprofit_margin（毛利率）错误：100% → 38.09% ✅
- **关键指标计算修复**:
  - 修复 equity_yoy（净资产同比增长率）字段映射错误
  - 添加 rd_exp_ratio（研发费用率）自动计算（从利润表）
  - 添加 ebit_to_interest（利息保障倍数）自动计算
- **数据表统一**: 统一使用带下划线的表名（balance_sheet、income_statement、cashflow）
- **代码清理**: 删除30+个临时测试/诊断脚本，优化项目结构

## 🆕 1.3.26版本功能

## 🆕 1.1版本新增功能

- **历史数据查询**: 支持查询行业阈值、企业评分、趋势分析的历史数据
- **趋势图标题优化**: 趋势图表标题更加简洁，只保留具体指标名称
- **系统名称更新**: 系统正式命名为"企业财务信息分析评价系统1.1版本"

## 🆕 最新优化功能（2026-01-04）

### 财务指标计算优化
- **指标详情功能**: 新增"📝 指标详情"选项卡，显示资产负债表字段详情、利润表字段详情、指标-字段对应关系
- **字段匹配优化**: 改进了字段匹配逻辑，使用匈牙利算法进行一对一匹配，确保每个PDF字段只能隶属于一个数据库字段
- **特殊匹配规则**: 添加了更多变体（"营业收入合计"、"净亏损"、"四、净利润"等），提高匹配准确性
- **字段预处理优化**: 改进了 `preprocess_account_name` 函数，使其更通用地处理序号、加减号、括号等特殊情况
- **相似度匹配**: 改进了 `is_special_match` 函数，使用相似度匹配（至少80%相似度）而不是精确匹配，提高匹配准确性和容错性

### PDF/图片处理优化
- **期初数据过滤**: 添加了期初数据过滤功能，当用户选择不识别期初数据时，期初字段不会进入匹配环节
- **表头识别优化**: 优化了表头识别逻辑，提高表头识别准确度：
  - 检查是否包含数值（如果包含数值，很可能是数据行而不是表头行）
  - 检查是否是有效的表头行（包含期间关键词、表头关键词、典型的表头词汇等）
  - 跳过包含数值的行（很可能是数据行）
  - 只保留真正的表头行
- **上传结果展示优化**: 改进了上传页面的上传结果展示逻辑：
  - 从数据库读取实际存储的数据，而不是直接显示PDF匹配结果
  - 显示各报表（资产负债表、利润表、现金流量表）的数据库存储数据
  - 在显示的时候把空值和none值都过滤掉不显示，仅显示有值的字段
  - 使用表格形式展示数据，一行英文字段、一行字段说明、一行具体值，看起来直观
  - 显示关键字段摘要
  - 保留字段匹配详情（PDF识别结果）作为参考信息

### 数据库存储优化
- **财务指标保存**: 添加了完整的财务指标保存函数 `save_financial_indicators_to_db`，支持重复记录处理（跳过、覆盖、更新）
- **数据库表字段问题修复**: 不手动插入 `created_at` 和 `updated_at` 列，让数据库自动处理这些时间戳字段
- **重复记录处理**: 使用 `_insert_single_record` 函数替代 `_insert_financial_ratios_record` 函数，确保正确处理重复记录

### 数据结构优化
- **数据结构定义**: 创建了 `financial_data_structures.py` 文件，包含数据结构定义：
  - `FinancialItem` 类：财务科目数据结构
  - `SheetExtractionResult` 类：报表提取结果

### 代码质量优化
- **缩进错误修复**: 修复了 `financial_indicators_page.py` 文件末尾的缩进错误
- **参数传递优化**: 修复了 `pdf_to_json_processor.py` 和 `streamlit_financial_pipeline_optimized.py` 中 `map_and_insert_to_database` 方法调用缺少 `include_beginning_period` 参数的问题

## 核心优化功能

- **行业阈值计算优化**: 行业阈值计算包含具体行业名称信息
- **评分计算优化**: 评分计算支持选择具体行业与企业匹配
- **行业自定义上传**: 支持上传自定义行业上市公司名单，生成行业阈值
- **多行业管理**: 支持Tushare标准行业和自定义行业的统一管理

## 🚀 快速开始

> 💡 **首次安装？** 请查看 [**📖 详细安装部署指南**](docs/安装部署指南.md) | **5分钟快速部署**

### 快速安装（5分钟）

```bash
# 1. 克隆代码
git clone https://github.com/daiyipu/EFAES.git
cd EFAES

# 2. 安装依赖
pip install -r requirements.txt

# 3. 配置数据库（编辑 src/web/config.py）
# 4. 初始化数据库和用户系统
python -c "from src.core.user_auth import init_auth_system; init_auth_system()"
python scripts/migrate_add_user_id.py

# 5. 启动应用
python scripts/start_financial_app_unified.py
```

👉 **详细安装步骤、常见问题解答：** [安装部署指南](docs/安装部署指南.md)

---

### 统一启动入口

项目提供统一的启动入口，简化启动流程：

```bash
# 启动财务数据处理系统（统一入口）
python scripts/start_financial_app_unified.py
```

### 公网访问启动

如果需要支持局域网或公网访问，可以使用以下命令：

```bash
# 启动支持公网访问的财务数据处理系统
cd scripts
python start_public_app.py
```

启动后，您可以通过以下地址访问应用：
- **本地访问**: http://localhost:8503
- **局域网访问**: http://[您的IP地址]:8503
- **公网访问**: 需要配置路由器端口转发或使用内网穿透工具

### 方法2：使用虚拟环境（最佳实践）

```bash
# 创建虚拟环境
python -m venv financial_env

# 激活虚拟环境
# Linux/Mac:
source financial_env/bin/activate
# Windows:
# financial_env\Scripts\activate

# 安装依赖
pip install -r requirements.txt
# 启动应用
cd scripts
python start_financial_app.py
```

### 方法3：使用bash脚本启动公网访问（推荐）
```bash
# 启动支持公网访问的财务数据处理系统
chmod +x start_public_streamlit.sh
./start_public_streamlit.sh
```


## 📊 应用访问地址

- **Streamlit应用**: http://localhost:8502
- **应用将在浏览器中自动打开**

## 🔧 问题解决

### 模块导入问题：No module named 'src'

**问题原因**：Python无法找到项目根目录中的 `src` 文件夹，因为项目根目录不在Python路径中。

**解决方案**：

```bash
# ✅ 正确方式：使用统一启动脚本（推荐）
cd scripts
python start_financial_app.py

# ✅ 方法2：手动设置Python路径
python setup_python_path.py
python -m src.web.streamlit_financial_pipeline_tushare_optimized

# ✅ 方法3：测试模块导入
python test_module_imports.py

# ❌ 错误方式：直接运行模块文件（会导致导入错误）
python src/web/streamlit_financial_pipeline_tushare_optimized.py
```

**重要提醒**：
- 项目采用标准Python包结构，`src` 是源代码目录
- 必须通过启动脚本或设置Python路径来访问 `src` 模块
- 直接运行模块文件会导致 "No module named 'src'" 错误

详细解决方案请参考：[`MODULE_IMPORT_SOLUTION.md`](MODULE_IMPORT_SOLUTION.md)

### 依赖版本冲突

如果遇到依赖版本冲突，请使用兼容版本：

```bash
# 卸载冲突版本
pip uninstall numpy huggingface-hub urllib3 fsspec

# 安装兼容版本
pip install -r requirements.txt
```

## 📋 功能特性

### 核心功能
- 📤 财务报表上传和处理
- 💾 数据存储和结构化
- 📊 行业分位数阈值计算（含具体行业名称）
- 🎯 企业财务评分计算（支持行业匹配）
- 📈 多年趋势分析可视化
- 🏭 行业自定义上传和管理
- 📚 历史数据查询（阈值、评分、趋势分析）

### 支持的界面
- **Streamlit界面**: 完整的财务数据处理流程（1.1版本）
- **Tushare集成**: 支持从Tushare获取实时行业数据
- **自定义行业**: 支持上传自定义行业上市公司名单
- **历史查询**: 支持查询数据库中的历史数据

## 🏗️ 项目结构

```
EFAES/                                          # 企业财务信息分析评价系统
├── src/                                        # 源代码目录
│   ├── core/                                  # 核心业务逻辑 (27个模块)
│   │   ├── calculate_financial_ratios.py      # 财务比率计算
│   │   ├── calculate_financial_score.py       # 财务评分计算
│   │   ├── finance_sheet_process.py           # 财务报表处理
│   │   ├── database_storage.py                # 数据库存储与字段匹配
│   │   ├── vision_financial_processor_optimized.py  # 视觉识别处理器
│   │   ├── pdf_to_json_processor.py           # PDF转JSON处理器
│   │   ├── industry_threshold_calculator.py   # 行业阈值计算
│   │   ├── trend_analysis_visualizer.py       # 趋势分析可视化
│   │   ├── financial_data_structures.py       # 数据结构定义
│   │   └── multi_model_vision_client.py       # 多模型视觉客户端
│   ├── web/                                   # Web界面 (10个模块)
│   │   ├── streamlit_financial_pipeline_optimized.py       # 主应用入口(推荐)
│   │   ├── upload_page.py                    # 文件上传页面
│   │   ├── financial_indicators_page.py       # 财务指标页面
│   │   ├── industry_threshold_page.py         # 行业阈值页面
│   │   ├── scoring_page.py                    # 评分计算页面
│   │   └── trend_analysis_page.py             # 趋势分析页面
│   ├── api/                                   # API接口
│   └── utils/                                 # 工具函数
├── scripts/                                   # 启动脚本
│   ├── start_financial_app.py                # 统一启动脚本
│   └── start_public_app.py                   # 公网访问启动脚本
├── data/                                      # 数据文件目录
│   ├── company_cache.pkl                     # 企业名称缓存
│   └── custom_industries.json                # 自定义行业数据
├── db/                                        # 数据库脚本
│   ├── create_financial_ratios_table.py      # 创建财务指标表
│   └── create_industry_thresholds_table.py   # 创建行业阈值表
├── docs/                                      # 文档目录
│   ├── API_USAGE_GUIDE.md                    # API使用指南
│   ├── ARCHITECTURE.md                       # 系统架构文档
│   └── CONTRIBUTING.md                       # 开发贡献指南
├── examples/                                  # 示例代码
├── logs/                                      # 日志文件
├── config/                                    # 配置文件
├── requirements.txt                           # 原始依赖
├── requirements.txt               # 兼容版本依赖
├── CHANGELOG.md                               # 版本变更记录
└── README.md                                  # 项目说明
```

## 📖 详细文档

- [系统架构文档](docs/ARCHITECTURE.md) - 了解系统整体架构和模块关系
- [API使用指南](docs/API_USAGE_GUIDE.md) - API接口详细说明
- [开发贡献指南](docs/CONTRIBUTING.md) - 如何参与项目开发
- [版本变更记录](CHANGELOG.md) - 历史版本更新内容
- [模块导入问题解决方案](docs/MODULE_IMPORT_SOLUTION.md) - 解决"No module named 'src'"错误

## 🎯 评分维度

### 1. 运营能力 (30%权重)
- 存货周转率
- 应收账款周转率
- 流动资产周转率
- 总资产周转率

### 2. 盈利能力 (30%权重)  
- 销售净利率
- 毛利率
- 净资产收益率
- 总资产收益率

### 3. 偿债能力 (20%权重)
- 流动比率
- 速动比率
- 资产负债率
- 利息保障倍数

### 4. 成长能力 (20%权重)
- 净利润同比增长率
- 营业收入同比增长率
- 净资产同比增长率

## 📊 评分等级

- **优秀**: 90-100分
- **良好**: 80-89分  
- **中等**: 70-79分
- **较差**: 60-69分
- **极差**: 0-59分

## 🔄 开发说明

### 使用自定义行业

1. 在Streamlit应用的"行业阈值计算"步骤中，选择"自定义新行业"标签页
2. 输入行业名称和上市公司股票代码列表
3. 系统会自动从Tushare获取这些公司的财务数据并生成行业阈值
4. 自定义行业数据会永久保存到 `custom_industries.json` 文件中

### 行业匹配评分

- 计算企业评分时，需要选择与企业行业匹配的行业阈值
- 支持Tushare标准行业和自定义行业的统一管理
- 行业阈值数据包含具体的行业名称信息

### 扩展行业支持

在相关模块中修改行业阈值字典，添加新的行业及其阈值标准。

### 扩展评分维度

修改财务评分计算模块，添加新的指标计算逻辑。

### 添加新的数据源

实现新的数据处理器类，支持不同的财务报表格式。

## ⚠️ 注意事项

1. 当前版本使用示例数据进行演示，实际应用中需要连接真实数据源
2. 行业阈值标准是示例数据，实际应用中应根据行业统计数据进行调整
3. 建议在生产环境中使用虚拟环境隔离依赖
4. 定期更新依赖包以获取安全更新和功能改进
5. 自定义行业功能需要有效的Tushare Token来获取上市公司财务数据
6. 行业匹配评分需要确保选择的行业与企业实际行业相符

## 📄 许可证

MIT License