# 定增盈利预测模型 — ML训练模块

基于定增/询转项目的多维特征，训练机器学习模型预测锁定期后的盈利概率。

当前最优模型：**LightGBM（类平衡）AUC=0.746, KS=0.363**（训练集 5折CV）。

## 目录结构

```
ml_training/
├── export_features.py              ① 特征提取
├── derive_features.py              ② 特征衍生
├── train_models.py                 ③ 双模型训练(LightGBM+逻辑回归)
├── train_profitability_model.py    ④ 轻量独立训练(仅Excel特征)
├── validate_model.py               ⑤ 外部测试集验证
│
├── data/                           特征数据
│   ├── features.parquet            基础特征 (1236行 × 277列)
│   ├── features.csv               同上CSV格式(查看用)
│   ├── features_derived.parquet   基础+衍生特征 (1236行 × 335列)
│   ├── features_derived.csv       同上CSV格式
│   └── features_schema.md         字段说明文档
│
└── output/                         模型与结果
    ├── lgb_classifier.txt          LightGBM分类模型
    ├── lr_classifier.pkl           逻辑回归模型(含scaler)
    ├── lr_coefficients.csv         逻辑回归系数表(评分卡)
    ├── lgb_feature_importance.csv  特征重要性数据
    ├── lgb_feature_importance.png  特征重要性图
    ├── evaluation_report.txt       评估报告
    ├── industry_analysis.csv       行业分析
    └── validation_result.csv       外部验证结果
```

## 执行顺序

```bash
# Step 1: 从scored Excel + MySQL 提取基础特征
python ml_training/export_features.py <scored.xlsx>

# Step 2: 在基础特征上派生高级特征
python ml_training/derive_features.py

# Step 3: 训练双模型
python ml_training/train_models.py data/features_derived.parquet --threshold -10

# Step 4: 外部测试集验证
python ml_training/validate_model.py <test.xlsx>
```

## 脚本详解

### ① export_features.py — 特征提取

**输入**：scored Excel（定增筛选结果）+ MySQL 数据库

**输出**：`data/features.parquet`（1236行 × 277列）

从 5 个数据源提取基础特征，以 `(股票代码, 报价日)` 为组合索引：

| 数据源 | 提取内容 | 来源 |
|--------|---------|------|
| market_data | 波动率(4窗口)、均线(5条)、胜率、漂移率 | investment_valuation MySQL |
| relative_valuation | 个股/行业 PE/PB/PS | investment_valuation MySQL |
| historical_fcf | 营收/NOPAT/FCF 等(最近5年) | investment_valuation MySQL |
| industry_data | 行业波动率、收益、胜率 | investment_valuation MySQL |
| peer_companies | 同行PE/PB/PS均值/中位 | investment_valuation MySQL |
| placement_params | 融资金额、锁定期、Beta | investment_valuation MySQL |
| screening_results | 筛选决策、子场景通过状态 | investment_valuation MySQL |
| financial_ratios | 28个财务比率 | fund_risk_control MySQL |
| scored Excel | 财务评分、子场景、7个月涨跌幅(标签) | Excel |

**关键机制**：
- 同一股票多次定增时，用报价日精确匹配对应时间点的数据
- 报价日与DB不匹配时，从 `price_series` 回算行情特征
- 16只股票有多次定增，97条样本的行情特征是回算的

### ② derive_features.py — 特征衍生

**输入**：`features.parquet` + MySQL（行业日线、大盘指数）

**输出**：`data/features_derived.parquet`（1236行 × 335列，新增58个衍生特征）

分两阶段执行，Stage 2 可通过 `--no-db` 跳过：

**Stage 1 — 纯文件运算（不访问数据库）**

| 类别 | 衍生方式 | 新特征 | 覆盖率 |
|------|---------|-------|--------|
| A. FCF增长率 | YoY / 2年CAGR / 增长加速度 | 18 | ~85% |
| B. FCF交叉比率 | FCF/营收、FCF/NOPAT、capex/折旧 | 5 | ~85% |
| C. 评分变动 | T-T-1 / T-T-4 年度delta | 9 | ~96% |
| D. 估值相对 | 个股PE/行业PE、PB/同行PB等 | 7 | ~98% |
| E. 行情动量 | 波动率比值、价格vs均线位置 | 7 | ~95% |

**Stage 2 — 需 MySQL**

| 类别 | 数据源 | 新特征 | 覆盖率 |
|------|--------|-------|--------|
| F. 行业PE/PB增长 | industry_daily (报价日前60/120/250天) | 6 | ~34% |
| G. 大盘指数特征 | market_indices (报价日时的大盘状态) | 6 | ~100% |

### ③ train_models.py — 双模型训练

**输入**：features.parquet 或 features_derived.parquet

训练两个对比模型，5折交叉验证：

| 模型 | AUC | 优势 | 输出文件 |
|------|-----|------|---------|
| LightGBM | 0.746 | 精准度高，捕捉非线性 | lgb_classifier.txt |
| 逻辑回归 | 0.558 | 可解释性强，系数=评分卡 | lr_classifier.pkl |

参数：
- `--threshold`：盈利阈值（默认-10%）
- `--threshold 0`：盈利>0%
- `--threshold -20`：盈利>-20%

### ④ train_profitability_model.py — 轻量独立训练

**输入**：scored Excel（直接读取）

仅从 Excel 自身的 14 个特征（评分+子场景）训练 LightGBM 分类+回归模型。
用于快速验证，不依赖数据库。AUC 约 0.50（接近随机），说明仅靠评分特征预测力不足。

### ⑤ validate_model.py — 外部测试集验证

**输入**：外部 Excel（如询转项目表）+ MySQL 数据库

完整流程：
1. 解析测试 Excel，构造与训练集相同的格式
2. 复用 export_features 提取 DB 特征
3. 复用 derive_features 计算衍生特征
4. 缺失特征用训练集 median 填充
5. 预测并与实际解禁浮盈对比

输出 AUC / KS / 混淆矩阵 / 概率分位分组表。

**注意**：询转项目特征覆盖率较低（~14%），大量特征用中位数填充，验证结果仅供参考。正式评估需先将测试集跑完整的数据采集流程。

## 模型性能

### 训练集（1236条定增样本，5折CV）

| 指标 | 基础特征(277列) | +衍生特征(335列) |
|------|---------------|----------------|
| AUC | 0.621 | **0.746** |
| KS | 0.204 | **0.363** |

### 外部验证（401条询转项目）

| 指标 | 盈利>0 | 盈利>-10% |
|------|--------|----------|
| AUC | 0.685 | **0.712** |
| KS | 0.341 | **0.437** |

（注意：测试集特征覆盖率仅14%，结果可信度有限）

## 关键发现

### 最重要的特征（LightGBM Top 5）

1. **市场波动率_120d** — 报价日时大盘的波动率，高波动率=牛市环境=盈利概率高
2. **市场波动率比值** — 短期vs长期波动率，衡量市场活跃度变化
3. **price_vs_MA20** — 股价在短期均线上的位置
4. **盈利能力_delta_1y** — 财务评分的年度改善幅度
5. **总分_delta_1y** — 综合评分的年度变化

### 核心结论

1. **择时最关键**：大盘在 MA250 上方时定增盈利率 47.5%，下方时仅 38.8%
2. **衍生特征贡献巨大**：58个衍生特征中有11个进入 Top 30
3. **市场特征 > 个股特征**：大盘环境对定增盈利的解释力最强
4. **财务评分有信号**：评分改善(delta_1y)比绝对值更有预测力

## 数据源依赖

| 数据库 | 地址 | 用途 |
|--------|------|------|
| investment_valuation | 127.0.0.1:3306 | 行情/估值/FCF/行业/同行/筛选结果 |
| fund_risk_control | 127.0.0.1:3306 | 财务比率/三表(覆盖极低) |

## 依赖库

- Python ≥ 3.10
- lightgbm, scikit-learn, pandas, numpy, pymysql, openpyxl
