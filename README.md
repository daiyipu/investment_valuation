# 投资估值系统

股权投资基金完整估值体系，涵盖相对估值、绝对估值、情景分析、压力测试等功能。

## 目录

- [系统概述](#系统概述)
- [系统架构](#系统架构)
- [模块设计](#模块设计)
- [技术栈](#技术栈)
- [快速开始](#快速开始)
- [使用指南](#使用指南)
- [API文档](#api文档)
- [开发指南](#开发指南)
- [项目路线图](#项目路线图)
- [贡献指南](#贡献指南)

---

## 系统概述

本系统为股权投资基金提供专业的投资估值工具，支持：

### 核心功能
- **相对估值法**：P/E、P/S、P/B、EV/EBITDA等多种倍数法
- **绝对估值法**：DCF现金流折现模型
- **其他估值方法**：VC法、成本法、交易对价法
- **情景分析**：基准/乐观/悲观多情景对比
- **压力测试**：收入冲击、毛利率压缩、WACC冲击、蒙特卡洛模拟
- **敏感性分析**：单因素/双因素分析、龙卷风图
- **数据获取**：集成Tushare API获取A股上市公司数据

### 适用场景
- 早期项目（天使轮、A轮）：VC法、交易对价法
- 成长期项目（B轮、C轮）：P/S法、DCF法
- 成熟期项目（Pre-IPO）：P/E法、DCF法、EV/EBITDA法

---

## 系统架构

### 整体架构图

```
┌─────────────────────────────────────────────────────────────────┐
│                         投资估值系统                             │
├─────────────────────────────────────────────────────────────────┤
│                                                                 │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐          │
│  │  Vue.js前端  │  │  FastAPI后端 │  │  Python核心  │          │
│  │              │  │              │  │   估值引擎   │          │
│  │ - Dashboard  │◄─┤ - RESTful API│◄─┤              │          │
│  │ - 估值输入   │  │ - CORS支持   │  │ - 相对估值   │          │
│  │ - 结果展示   │  │ - 数据验证   │  │ - 绝对估值   │          │
│  │ - 图表可视化 │  │              │  │ - 情景分析   │          │
│  └──────────────┘  └──────────────┘  │ - 压力测试   │          │
│                                  │  │ - 敏感性分析 │          │
│                                  │  └──────────────┘          │
│                                  │                             │
│                                  │  ┌──────────────┐          │
│                                  │  │ Tushare数据  │          │
│                                  │  │    获取器    │          │
│                                  │  └──────┬───────┘          │
│                                  │         │                  │
│                                  │         ▼                  │
│                                  │  ┌──────────────┐          │
│                                  │  │  Tushare API │          │
│                                  │  └──────────────┘          │
│                                  │                             │
└──────────────────────────────────┴─────────────────────────────┘
```

### 分层架构

```
┌─────────────────────────────────────────────────────────────┐
│                      表现层 (Presentation)                    │
│  Vue.js 前端 / API 文档 / 命令行工具                         │
├─────────────────────────────────────────────────────────────┤
│                      API层 (API Gateway)                     │
│  FastAPI / RESTful API / 数据验证 / CORS                    │
├─────────────────────────────────────────────────────────────┤
│                      业务层 (Business Logic)                  │
│  ┌──────────┬──────────┬──────────┬──────────┬──────────┐   │
│  │ 相对估值 │ 绝对估值 │ 情景分析 │ 压力测试 │ 敏感性   │   │
│  └──────────┴──────────┴──────────┴──────────┴──────────┘   │
├─────────────────────────────────────────────────────────────┤
│                      数据层 (Data Access)                     │
│  Tushare API / 本地计算 / 内存缓存                           │
├─────────────────────────────────────────────────────────────┤
│                      模型层 (Domain Models)                   │
│  Company / Comparable / ValuationResult / Scenario           │
└─────────────────────────────────────────────────────────────┘
```

### 数据流图

```
┌─────────┐    ┌─────────┐    ┌─────────┐    ┌─────────┐
│  用户   │───►│  前端   │───►│  API   │───►│  估值   │
│  输入   │    │  表单   │    │  网关   │    │  引擎   │
└─────────┘    └─────────┘    └─────────┘    └────┬────┘
                                                  │
              ┌───────────────────────────────────┘
              │
              ▼
    ┌─────────────────┐
    │   估值计算      │
    │   - 参数验证    │
    │   - 模型选择    │
    │   - 结果计算    │
    └────────┬────────┘
             │
             ▼
    ┌─────────────────┐
    │   结果处理      │
    │   - 格式化      │
    │   - 汇总        │
    │   - 报告生成    │
    └────────┬────────┘
             │
             ▼
    ┌─────────────────┐
    │   返回用户      │
    │   - JSON数据    │
    │   - 可视化图表  │
    │   - 分析报告    │
    └─────────────────┘
```

---

## 模块设计

### 1. 核心数据模型 (models.py)

```python
Company          # 公司实体模型
Comparable       # 可比公司模型
ValuationResult  # 估值结果容器
ScenarioConfig   # 情景配置
StressTestResult # 压力测试结果
MonteCarloResult # 蒙特卡洛结果
```

**职责**：
- 定义领域对象
- 数据验证和转换
- 提供数据结构化接口

### 2. 相对估值模块 (relative_valuation.py)

```python
RelativeValuation
├── pe_ratio_valuation()        # 市盈率法
├── ps_ratio_valuation()        # 市销率法
├── pb_ratio_valuation()        # 市净率法
├── ev_ebitda_valuation()       # EV/EBITDA法
└── auto_comparable_analysis()  # 自动可比分析
```

**职责**：
- 实现各种相对估值算法
- 处理可比公司数据
- 计算估值倍数和区间

### 3. 绝对估值模块 (absolute_valuation.py)

```python
AbsoluteValuation
├── calculate_wacc()              # WACC计算
├── calculate_terminal_value()    # 终值计算
├── forecast_free_cash_flows()    # 现金流预测
├── dcf_valuation()               # DCF估值
└── dcf_sensitivity_analysis()    # DCF敏感性分析
```

**职责**：
- 实现DCF估值模型
- 计算加权平均资本成本
- 预测未来现金流
- 终值计算（永续增长/退出倍数）

### 4. 情景分析模块 (scenario_analysis.py)

```python
ScenarioAnalyzer
├── base_case()               # 基准情景
├── bull_case()               # 乐观情景
├── bear_case()               # 悲观情景
├── custom_scenario()         # 自定义情景
├── compare_scenarios()       # 多情景对比
└── scenario_probability_analysis()  # 概率分析
```

**职责**：
- 管理多种情景配置
- 执行情景估值计算
- 生成情景对比报告

### 5. 压力测试模块 (stress_test.py)

```python
StressTester
├── revenue_shock_test()        # 收入冲击测试
├── margin_compression_test()   # 利润率压缩测试
├── wacc_shock_test()           # WACC冲击测试
├── growth_slowdown_test()      # 增长放缓测试
├── extreme_market_crash()      # 极端情景测试
├── monte_carlo_simulation()    # 蒙特卡洛模拟
└── generate_stress_report()    # 综合压力报告
```

**职责**：
- 执行各类压力测试
- 蒙特卡洛随机模拟
- 生成压力测试报告

### 6. 敏感性分析模块 (sensitivity_analysis.py)

```python
SensitivityAnalyzer
├── one_way_sensitivity()        # 单因素敏感性
├── two_way_sensitivity()        # 双因素敏感性
├── tornado_chart_data()         # 龙卷风图数据
└── comprehensive_sensitivity()  # 综合敏感性
```

**职责**：
- 单因素/双因素敏感性计算
- 生成龙卷风图数据
- 参数影响排序

### 7. 数据获取模块 (data_fetcher.py)

```python
TushareDataFetcher
├── get_comparable_companies()  # 获取可比公司
├── get_financial_metrics()     # 获取财务指标
├── get_industry_multiples()    # 获取行业倍数
└── search_by_keywords()        # 关键词搜索
```

**职责**：
- 从Tushare获取上市公司数据
- 提取财务指标和估值倍数
- 行业数据聚合

### 8. 其他估值方法模块 (other_methods.py)

```python
OtherValuationMethods
├── vc_method()                      # VC法
├── cost_method()                    # 成本法
├── transaction_comparable()         # 交易对价法
├── first_chicago_method()           # 第一芝加哥法
└── sum_of_parts_valuation()         # 分部加总法
```

**职责**：
- 实现特殊估值方法
- 早期项目估值
- 资产重组估值

### 9. API服务模块 (api.py)

```python
FastAPI Application
├── /api/valuation/*     # 估值相关API
├── /api/scenario/*      # 情景分析API
├── /api/stress-test/*   # 压力测试API
└── /api/sensitivity/*   # 敏感性分析API
```

**职责**：
- 提供RESTful API接口
- 请求数据验证
- 响应格式化
- CORS跨域支持

---

## 技术栈

### 后端
- **Python 3.8+**
- **FastAPI**: 现代化Web框架
- **Pydantic**: 数据验证
- **NumPy**: 数值计算
- **Pandas**: 数据处理

### 数据源
- **Tushare**: A股数据接口

### 前端（待开发）
- **Vue.js 3**: 前端框架
- **Vue Router**: 路由管理
- **ECharts**: 图表可视化
- **Element Plus**: UI组件库
- **Axios**: HTTP客户端

---

## 快速开始

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行示例

```bash
python examples.py
```

### 启动API服务

```bash
python api.py
```

访问 API 文档：http://localhost:8000/docs

---

## 使用指南

### 基本估值

```python
from models import Company, CompanyStage
from relative_valuation import RelativeValuation

# 创建公司
company = Company(
    name="测试公司",
    industry="软件服务",
    stage=CompanyStage.GROWTH,
    revenue=50000,
    net_income=8000,
    growth_rate=0.25,
)

# 相对估值
comparables = [...]  # 可比公司列表
results = RelativeValuation.auto_comparable_analysis(company, comparables)
```

### DCF估值

```python
from absolute_valuation import AbsoluteValuation

result = AbsoluteValuation.dcf_valuation(company, projection_years=5)
print(f"估值: {result.value/10000:.2f}亿元")
```

### 情景分析

```python
from scenario_analysis import ScenarioAnalyzer

analyzer = ScenarioAnalyzer(company)
results = analyzer.compare_scenarios()
```

### 压力测试

```python
from stress_test import StressTester

tester = StressTester(company)
report = tester.generate_stress_report()
```

---

## API文档

### 估值端点

#### POST /api/valuation/relative
相对估值法

#### POST /api/valuation/absolute
DCF绝对估值法

#### POST /api/valuation/compare
多种方法交叉验证

### 情景分析端点

#### POST /api/scenario/analyze
多情景分析

### 压力测试端点

#### POST /api/stress-test/revenue
收入冲击测试

#### POST /api/stress-test/monte-carlo
蒙特卡洛模拟

#### POST /api/stress-test/full
综合压力测试

### 敏感性分析端点

#### POST /api/sensitivity/one-way
单因素敏感性分析

#### POST /api/sensitivity/tornado
龙卷风图数据

#### POST /api/sensitivity/comprehensive
综合敏感性分析

详细API文档请访问：http://localhost:8000/docs

---

## 开发指南

### 项目结构

```
Investment_valuation/
├── models.py                  # 数据模型
├── data_fetcher.py            # 数据获取
├── relative_valuation.py      # 相对估值
├── absolute_valuation.py      # 绝对估值
├── other_methods.py           # 其他估值方法
├── scenario_analysis.py       # 情景分析
├── stress_test.py             # 压力测试
├── sensitivity_analysis.py    # 敏感性分析
├── api.py                     # FastAPI服务
├── examples.py                # 使用示例
├── requirements.txt           # 依赖包
└── README.md                  # 项目文档
```

### 代码规范

- 使用类型注解
- 函数添加文档字符串
- 遵循PEP 8规范
- 使用dataclass定义数据模型

### 测试

```bash
# 运行所有示例
python examples.py

# 运行单元测试（待开发）
pytest tests/
```

---

## 项目路线图

### 📊 当前版本：v1.0.0

**已完成功能**：
- ✅ DCF绝对估值
- ✅ 相对估值（PE/PS/PB/EV-EBITDA）
- ✅ 情景分析
- ✅ 压力测试
- ✅ 敏感性分析
- ✅ Vue.js前端界面
- ✅ Tushare数据集成
- ✅ 报告生成与导出

### 🚀 开发计划

详细的迭代规划请查看：[项目路线图](./PROJECT_ROADMAP.md)

- **v1.1** (进行中): 性能优化与用户体验提升
- **v1.2** (计划): 数据可视化增强
- **v1.3** (计划): 高级分析功能
- **v1.4** (计划): 系统集成与扩展
- **v1.5** (计划): 移动端适配
- **v2.0** (计划): 智能化升级（AI/ML）

### 📋 当前迭代

查看当前迭代（v1.1）的详细任务：[迭代任务分解](./SPRINT_1.1_TASKS.md)

---

## 贡献指南

我们欢迎所有形式的贡献！

### 🤝 如何贡献

1. 查看 [贡献指南](./CONTRIBUTING.md)
2. Fork本项目
3. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
4. 提交更改 (`git commit -m 'feat: add some AmazingFeature'`)
5. 推送到分支 (`git push origin feature/AmazingFeature`)
6. 提交Pull Request

### 📝 报告问题

- Bug报告：使用 [Bug报告模板](.github/ISSUE_TEMPLATE/bug_report.md)
- 功能建议：使用 [功能建议模板](.github/ISSUE_TEMPLATE/feature_request.md)
- 性能问题：使用 [性能问题模板](.github/ISSUE_TEMPLATE/performance.md)

### 🎯 代码规范

- **Python**: 遵循PEP 8
- **TypeScript/Vue**: 遵循Vue风格指南
- **提交信息**: 遵循[Conventional Commits](https://www.conventionalcommits.org/)

### ✅ 测试要求

- 单元测试覆盖率 ≥ 80%
- 所有测试通过才能合并
- 关键业务逻辑必须有测试

---

## CI/CD

本项目使用GitHub Actions进行持续集成和持续部署：

- ✅ 自动化测试
- ✅ 代码质量检查
- ✅ 安全扫描
- ✅ 性能测试
- ✅ 自动部署

查看详细配置：[CI/CD配置](.github/workflows/ci.yml)

---

## 许可证

MIT License

---

## 联系方式

- GitHub: https://github.com/daiyipu/investment_valuation
- Issues: https://github.com/daiyipu/investment_valuation/issues
- Discussions: https://github.com/daiyipu/investment_valuation/discussions

---

## 致谢

感谢所有为本项目做出贡献的开发者！

特别感谢：
- Tushare提供数据支持
- FastAPI团队提供优秀的框架
- Vue.js团队提供强大的前端框架

---

## ⭐ 如果这个项目对你有帮助，请给个Star！


