# 定增项目风险分析模块

## 概述

本模块专门用于私募股权投资中定增（定向增发）项目的风险分析，通过 Jupyter Notebook 进行交互式数据分析。

## 目录结构

```
price_maintenance_risk_analysis/
├── notebooks/              # Jupyter Notebook 分析文件
│   ├── 01_sensitivity_analysis.ipynb      # 敏感性分析
│   ├── 02_stress_test.ipynb              # 压力测试
│   ├── 03_dcf_valuation.ipynb            # DCF估值
│   ├── 04_monte_carlo_simulation.ipynb   # 蒙特卡洛模拟
│   ├── 05_var_calculation.ipynb          # VaR风险测算
│   └── 06_comprehensive_analysis.ipynb   # 综合分析报告
│
├── data/                   # 样例数据
│   └── sample_data.csv
│
├── utils/                  # 分析工具函数
│   └── analysis_tools.py
│
└── README.md              # 本文件
```

## 分析内容

### 1. 敏感性分析 (01_sensitivity_analysis.ipynb)
- 锁定期价格敏感性分析
- 破发风险测算
- 关键参数影响分析（发行折价率、解禁比例等）

### 2. 压力测试 (02_stress_test.ipynb)
- 极端市场情景测试
- 行业危机情景
- 个股重大利空情景
- 组合压力测试

### 3. DCF估值 (03_dcf_valuation.ipynb)
- 定增项目现金流预测
- 收益率测算（年化收益率、IRR）
- 估值安全边际分析
- 不同退出方式估值对比

### 4. 蒙特卡洛模拟 (04_monte_carlo_simulation.ipynb)
- 股价路径模拟（几何布朗运动）
- 定增收益分布模拟
- 概率分析（盈利概率、亏损概率）
- 收益率区间预测

### 5. VaR风险测算 (05_var_calculation.ipynb)
- 历史模拟法计算 VaR
- 方差-协方差法计算 VaR
- 蒙特卡洛法计算 VaR
- CVaR（条件风险价值）测算
- 最大回撤测算

### 6. 综合分析报告 (06_comprehensive_analysis.ipynb)
- 项目整体风险评估
- 收益风险比分析
- 投资建议生成

## 使用方法

### 环境配置

```bash
cd price_maintenance_risk_analysis

# 安装依赖
pip install jupyter notebook pandas numpy scipy matplotlib seaborn

# 启动 Jupyter
jupyter notebook
```

### 数据输入

在 `data/` 目录下准备以下数据：
- 项目基本信息（发行价格、发行数量、锁定期等）
- 标的股票历史价格数据
- 财务数据
- 行业对比数据

## 输出报告

每个分析 Notebook 会生成：
- 数据可视化图表
- 风险指标计算结果
- PDF 报告（可选）

## 故障排除

### 中文显示乱码

如果图表中的中文显示为方块或乱码，请使用以下方法解决：

#### 方法 1: 运行字体诊断工具

```bash
# 在项目目录下运行
python diagnose_font.py

# 或者测试字体
python diagnose_font.py test

# 清除 matplotlib 缓存
python diagnose_font.py clear-cache
```

#### 方法 2: 手动安装中文字体

**Linux 系统:**
```bash
# Ubuntu/Debian
sudo apt-get install fonts-noto-cjk fonts-wqy-microhei

# CentOS/RedHat
sudo yum install wqy-microhei-fonts wqy-zenhei-fonts

# 安装后清除缓存
fc-cache -fv
```

**macOS 系统:**
- 系统已内置中文字体（PingFang SC、华文黑体）
- 如有问题，打开"字体册"应用检查字体状态

**Windows 系统:**
- 系统已内置微软雅黑字体
- 确认 `C:\Windows\Fonts\` 目录下有 `msyh.ttc`

#### 方法 3: 运行字体安装助手

```bash
python install_chinese_font.py
```

#### vnpy 环境

vnpy 环境使用独立的 Python 解释器，可能需要额外配置：

1. 激活 vnpy 虚拟环境后安装字体
2. 在 notebook 第一个代码块中运行：
   ```python
   from utils.font_config import setup_chinese_font
   setup_chinese_font()
   ```
3. 清除 matplotlib 缓存并重启 Jupyter
