# 定增风险分析系统

## 目录结构

```
.
├── data/                           # 数据文件
│   ├── 300735_SZ_placement_params.json  # 定增参数
│   ├── 300735_SZ_market_data.json       # 市场数据
│   └── sample_stock_data.csv             # 示例数据
│
├── notebooks/                      # Jupyter 分析笔记本
│   ├── 01_sensitivity_analysis.ipynb     # 敏感性分析
│   ├── 02_stress_test.ipynb             # 压力测试
│   ├── 03_dcf_valuation.ipynb           # DCF估值
│   ├── 04_monte_carlo_simulation.ipynb  # 蒙特卡洛模拟
│   ├── 05_var_calculation.ipynb         # VaR计算
│   ├── 06_comprehensive_analysis.ipynb  # 综合分析
│   └── 07_market_data_analysis.ipynb    # 市场数据分析
│
├── reports/                        # 报告输出
│   ├── images/                        # 分析图表
│   └── outputs/                       # 生成的报告文件
│
├── scripts/                        # 工具脚本
│   ├── generate_word_report.py         # Word报告生成器
│   ├── fetch_gh_data.py                # 数据获取脚本
│   └── update_market_data.py           # 市场数据更新
│
└── utils/                          # 工具模块
    ├── analysis_tools.py              # 分析工具
    ├── config_loader.py               # 配置加载器
    ├── market_data_loader.py          # 市场数据加载
    └── font_manager.py                # 字体管理
```

## 使用说明

### 1. 数据获取

```bash
# 更新市场数据
cd scripts
python update_market_data.py
```

### 2. 运行分析

在 Jupyter Notebook 中按顺序运行：
- `01_sensitivity_analysis.ipynb`
- `02_stress_test.ipynb`
- `03_dcf_valuation.ipynb`
- `04_monte_carlo_simulation.ipynb`
- `05_var_calculation.ipynb`
- `06_comprehensive_analysis.ipynb`

### 3. 生成报告

```bash
cd scripts
python generate_word_report.py
```

报告将生成在 `reports/outputs/` 目录下。
