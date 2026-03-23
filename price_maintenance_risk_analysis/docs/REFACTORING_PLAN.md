# generate_word_report_v2.py 模块化重构方案

## 📋 概述

**目标：** 将 `generate_word_report_v2.py`（9368行）按章节拆分成多个子模块，降低单个文件复杂度

**状态：** 🟡 框架已搭建，待实施

**创建时间：** 2025-03-23

---

## 📁 目录结构

```
scripts/generate_word_report_v2/
├── __init__.py              # 包初始化，导出 generate_report
├── main.py                  # ✅ 统一入口脚本（已创建）
├── utils.py                 # ✅ 通用工具函数（已创建框架）
├── chapter01_overview.py    # 🟡 项目概况（框架已创建）
├── chapter02_valuation.py   # 🟡 相对估值分析（框架已创建）
├── chapter03_dcf.py         # 🟡 DCF估值分析（框架已创建）
├── chapter04_sensitivity.py # 🟡 敏感性分析（框架已创建）
├── chapter05_montecarlo.py  # 🟡 蒙特卡洛模拟（框架已创建）
├── chapter06_scenario.py    # 🟡 情景分析（框架已创建）
├── chapter07_stress.py      # 🟡 压力测试（框架已创建）
├── chapter08_var.py         # 🟡 VaR风险度量（框架已创建）
├── chapter09_advice.py      # 🟡 风控建议与风险提示（框架已创建）
└── appendix_scenarios.py    # 🟡 附件：情景数据表（框架已创建）
```

---

## 📊 原文件结构分析

### 文件大小
- **总行数：** 9368行
- **章节数量：** 10个主要章节 + 附件

### 章节分布

| 章节 | 行号范围 | 说明 |
|-----|---------|------|
| 封面+目录 | 2768-2782 | 报告封面和目录 |
| 第一章：项目概况 | 2783-3238 | 市场数据、项目参数 |
| 第二章：相对估值分析 | 3239-3726 | PE/PB/PS对比、PE历史分位数 |
| 第三章：DCF估值分析 | 3727-4395 | DCF模型、敏感性分析 |
| 第四章：敏感性分析 | 4395-5603 | 发行价折扣、参数敏感性 |
| 第五章：蒙特卡洛模拟 | 5604-6547 | 基础分析、ARIMA/GARCH预测 |
| 第六章：情景分析 | 6548-7487 | 多参数情景、PE分位数情景 |
| 第七章：压力测试 | 7488-8028 | 极端情景分析 |
| 第八章：VaR风险度量 | 8029-8502 | VaR计算、CVaR |
| 第九章：风控建议 | 8844-9212 | 综合评估、风险提示 |
| 附件：情景数据表 | 9213-9368 | 585种情景完整数据 |

---

## 🔧 核心设计

### 1. Context 数据流

```python
context = {
    # 基础配置
    'stock_code': '300735.SZ',
    'stock_name': '光弘科技',

    # 数据源
    'project_params': {...},  # 定增项目参数
    'market_data': {...},     # 市场数据
    'industry_data': {...},   # 行业数据

    # 工具对象
    'document': Document(),   # Word文档对象
    'analyzer': Analyzer(),   # 风险分析器
    'font_prop': FontProp(),  # 字体属性

    # 路径配置
    'IMAGES_DIR': '...',      # 图片保存目录
    'DATA_DIR': '...',        # 数据目录
    'OUTPUTS_DIR': '...',     # 输出目录

    # 计算结果（章节间传递）
    'results': {
        'mc_results': {...},      # 蒙特卡洛模拟结果
        'scenarios_config': [...], # 情景配置
        'var_results': {...},     # VaR计算结果
        # ... 其他计算结果
    }
}
```

### 2. 章节模块接口

```python
def generate_chapter(context):
    """
    生成章节内容

    参数:
        context: 包含所有必要数据的字典
    """
    document = context['document']
    project_params = context['project_params']

    # 使用 utils 模块的辅助函数
    from . import utils
    utils.add_title(document, '章节标题', level=1)
    utils.add_paragraph(document, '章节内容')

    # 计算结果保存到 context
    context['results']['chapter_result'] = {...}
```

### 3. 主流程控制

```python
# main.py
def generate_report(stock_code, stock_name):
    document = Document()
    context = {...}

    # 依次调用各章节
    chapter01_overview.generate_chapter(context)
    chapter02_valuation.generate_chapter(context)
    # ... 其他章节

    return document
```

---

## ✅ 已完成工作

### 1. 框架搭建
- ✅ 创建模块目录结构
- ✅ 创建 `__init__.py`
- ✅ 创建 `main.py` 统一入口
- ✅ 创建 `utils.py` 工具函数模块
- ✅ 创建10个章节占位符文件

### 2. 工具函数提取
- ✅ Word文档格式化函数（add_title、add_paragraph、add_table等）
- ⏳ 图表生成函数（框架已创建，待提取具体实现）

### 3. 文档编写
- ✅ 重构方案文档（本文档）
- ✅ 代码注释和TODO标注

---

## ⏳ 待完成工作

### 阶段1：提取工具函数实现（优先级：高）

**文件：** `utils.py`

**需要提取的函数：**
- `generate_break_even_chart()`
- `generate_lockup_sensitivity_charts_split()`
- `generate_time_window_analysis_chart()`
- `generate_tornado_chart_split()`
- `generate_discount_scenario_charts_split()`
- `generate_sensitivity_3d_charts_split()`
- `generate_monte_carlo_charts_split()`
- `generate_stress_test_charts_split()`
- `generate_var_charts_split()`
- `generate_relative_valuation_charts_split()`
- `generate_dcf_charts_split()`
- `generate_market_data_charts_split()`
- `generate_industry_index_charts_split()`

**预计工作量：** 2-3小时

---

### 阶段2：提取章节内容（优先级：高）

#### chapter01_overview.py - 项目概况
**源文件行号：** 2783-3238（约455行）

**需要提取的内容：**
- 封面生成
- 目录生成
- 基本信息展示
- 市场数据表格
- 行业数据表格
- 图表生成

#### chapter02_valuation.py - 相对估值分析
**源文件行号：** 3239-3726（约487行）

**需要提取的内容：**
- 2.1 估值指标对比
- 2.2 PE、PB、PS相对位置分析
- 2.3 PE历史分位数趋势分析
- 相对估值图表生成

#### chapter03_dcf.py - DCF估值分析
**源文件行号：** 3727-4395（约668行）

**需要提取的内容：**
- 3.1 DCF估值方法说明
- 3.2 基础数据
- 3.3 DCF计算
- 3.4 敏感性分析
- 3.5 估值结论

#### chapter04_sensitivity.py - 敏感性分析
**源文件行号：** 4395-5603（约1208行）

**需要提取的内容：**
- 4.1 盈亏平衡分析
- 4.2 发行价折扣敏感性
- 4.3 参数敏感性分析（龙卷风图）
- 4.4 三维敏感性分析

#### chapter05_montecarlo.py - 蒙特卡洛模拟
**源文件行号：** 5604-6547（约943行）

**需要提取的内容：**
- 5.1 蒙特卡洛模拟基础分析
- 5.2 多窗口期对比分析
- 5.3 ARIMA预测漂移率
- 5.4 GARCH预测波动率
- 5.5 基于预测参数的模拟

#### chapter06_scenario.py - 情景分析
**源文件行号：** 6548-7487（约939行）

**需要提取的内容：**
- 6.1 多参数情景分析
- 6.2 基于行业PE分位数的情景
- 6.3 基于个股PE分位数的情景
- 6.4 破发概率分析

#### chapter07_stress.py - 压力测试
**源文件行号：** 7488-8028（约540行）

**需要提取的内容：**
- 7.1 极端情景定义
- 7.2 压力测试结果
- 7.3 抗风险能力评估

#### chapter08_var.py - VaR风险度量
**源文件行号：** 8029-8502（约473行）

**需要提取的内容：**
- 8.1 VaR方法说明
- 8.2 VaR计算
- 8.3 CVaR分析
- 8.4 回测验证

#### chapter09_advice.py - 风控建议
**源文件行号：** 8844-9212（约368行）

**需要提取的内容：**
- 9.1 综合风险评估
- 9.2 风控建议
- 9.3 风险提示

#### appendix_scenarios.py - 附件
**源文件行号：** 9213-9368（约155行）

**需要提取的内容：**
- 585种情景完整数据表

**预计工作量：** 8-12小时

---

### 阶段3：测试与优化（优先级：中）

**测试任务：**
- [ ] 单元测试（各章节模块）
- [ ] 集成测试（完整报告生成）
- [ ] 性能测试（生成时间对比）
- [ ] 边界测试（数据缺失等异常情况）

**优化任务：**
- [ ] 代码注释完善
- [ ] 错误处理增强
- [ ] 日志输出优化
- [ ] 文档补充完整

**预计工作量：** 4-6小时

---

## 🚀 使用方法

### 当前状态（框架阶段）

```bash
# 方式1：使用原始单文件脚本（推荐）
cd scripts
python generate_word_report_v2.py

# 方式2：使用模块化版本（待完成）
cd scripts/generate_word_report_v2
python main.py --stock 300735.SZ --name 光弘科技
```

### 完成后（模块化版本）

```bash
# 基本使用
python -m generate_word_report_v2.main --stock 300735.SZ

# 自定义参数
python main.py \
    --stock 300735.SZ \
    --name 光弘科技 \
    --output my_report.docx

# 在Python中调用
from generate_word_report_v2 import generate_report
doc = generate_report('300735.SZ', '光弘科技')
doc.save('report.docx')
```

---

## 📈 预期收益

### 代码质量提升
- ✅ **可维护性：** 单个文件从9368行降至各章节500-1200行
- ✅ **可读性：** 章节职责清晰，易于理解
- ✅ **可扩展性：** 新增章节只需添加新模块

### 开发效率提升
- ✅ **并行开发：** 不同章节可独立开发
- ✅ **快速定位：** 问题快速定位到具体章节
- ✅ **复用性：** 工具函数可在其他项目复用

### 团队协作友好
- ✅ **代码审查：** 可按章节分批审查
- ✅ **任务分配：** 可按章节分配任务
- ✅ **版本管理：** 减少merge冲突

---

## ⚠️ 注意事项

### 1. 原文件保留
- 原始 `generate_word_report_v2.py` **保持不变**
- 新的模块化版本作为平行分支
- 待稳定后可考虑替换原文件

### 2. 向后兼容
- 确保生成的报告格式一致
- 确保所有功能完整保留
- 确保性能不低于原版本

### 3. 数据依赖
- 各章节通过 context 共享数据
- 注意数据依赖顺序（某些章节依赖前面的计算结果）
- 避免循环依赖

### 4. 错误处理
- 单个章节失败不应影响其他章节
- 需要增强异常捕获和错误提示
- 提供降级方案（如数据缺失时使用默认值）

---

## 📝 TODO 清单

### 高优先级
- [ ] 提取所有图表生成函数到 utils.py
- [ ] 完成 chapter01_overview.py 实现
- [ ] 完成 chapter02_valuation.py 实现
- [ ] 完成 chapter05_montecarlo.py 实现（ARIMA/GARCH部分）

### 中优先级
- [ ] 完成其余章节模块
- [ ] 编写单元测试
- [ ] 性能优化

### 低优先级
- [ ] 代码注释完善
- [ ] 文档补充
- [ ] 示例代码

---

## 📞 联系方式

如有疑问或建议，请联系项目维护者或提交Issue。

**项目地址：** https://github.com/[username]/investment-valuation

**最后更新：** 2025-03-23
