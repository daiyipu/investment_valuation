# 特征字段说明


## 基础信息 (7个)

| 字段 | 来源 |
|------|------|
| 股票代码 | SQLite |
| 股票简称 | Excel |
| 报价日 | SQLite |
| 报价日_excel | Excel |
| 一级行业 | Excel |
| 二级行业 | Excel |
| 三级行业 | Excel |

## 行情特征(SQLite) (27个)

| 字段 | 来源 |
|------|------|
| 波动率_20d | SQLite |
| 年化收益_20d | SQLite |
| 区间收益_20d | SQLite |
| 胜率_20d | SQLite |
| 波动率_60d | SQLite |
| 年化收益_60d | SQLite |
| 区间收益_60d | SQLite |
| 胜率_60d | SQLite |
| 波动率_120d | SQLite |
| 年化收益_120d | SQLite |
| 区间收益_120d | SQLite |
| 胜率_120d | SQLite |
| 波动率_250d | SQLite |
| 年化收益_250d | SQLite |
| 区间收益_250d | SQLite |
| 胜率_250d | SQLite |
| MA20 | SQLite |
| MA30 | SQLite |
| MA60 | SQLite |
| MA120 | SQLite |
| MA250 | SQLite |
| 当前价 | SQLite |
| 漂移率 | SQLite |
| 波动率 | SQLite |
| 换手率 | SQLite |
| 数据天数 | SQLite |
| 报价日MA20 | SQLite |

## 估值特征(SQLite) (8个)

| 字段 | 来源 |
|------|------|
| 个股PE | SQLite |
| 个股PB | SQLite |
| 个股PS | SQLite |
| 行业PE | SQLite |
| 行业PB | SQLite |
| 行业PS | SQLite |
| 行业代码 | SQLite |
| 行业名称 | SQLite |

## FCF特征(SQLite) (20个)

| 字段 | 来源 |
|------|------|
| 营收_T | SQLite |
| NOPAT_T | SQLite |
| FCF_T | SQLite |
| FCF年份_T | SQLite |
| 营收_T1 | SQLite |
| NOPAT_T1 | SQLite |
| FCF_T1 | SQLite |
| FCF年份_T1 | SQLite |
| 营收_T2 | SQLite |
| NOPAT_T2 | SQLite |
| FCF_T2 | SQLite |
| FCF年份_T2 | SQLite |
| 营收_T3 | SQLite |
| NOPAT_T3 | SQLite |
| FCF_T3 | SQLite |
| FCF年份_T3 | SQLite |
| 营收_T4 | SQLite |
| NOPAT_T4 | SQLite |
| FCF_T4 | SQLite |
| FCF年份_T4 | SQLite |

## 定增参数(SQLite) (5个)

| 字段 | 来源 |
|------|------|
| 发行价 | SQLite |
| 锁定期 | SQLite |
| 净债务 | SQLite |
| 净利润 | SQLite |
| 净资产负债表 | SQLite |

## 筛选决策(SQLite) (7个)

| 字段 | 来源 |
|------|------|
| 溢价率下限 | SQLite |
| 溢价率上限 | SQLite |
| 有效阈值数 | SQLite |
| step1通过 | SQLite |
| step2通过 | SQLite |
| step3通过 | SQLite |
| 定增决策 | SQLite |

## 财务评分(Excel) (27个)

| 字段 | 来源 |
|------|------|
| 总分_2021 | Excel |
| 总分_2022 | Excel |
| 总分_2023 | Excel |
| 总分_2024 | Excel |
| 总分_2025 | Excel |
| 评级_2021 | Excel |
| 评级_2022 | Excel |
| 评级_2023 | Excel |
| 评级_2024 | Excel |
| 评级_2025 | Excel |
| 盈利能力_2021 | Excel |
| 盈利能力_2022 | Excel |
| 盈利能力_2023 | Excel |
| 盈利能力_2024 | Excel |
| 盈利能力_2025 | Excel |
| 成长能力_2021 | Excel |
| 成长能力_2022 | Excel |
| 成长能力_2023 | Excel |
| 成长能力_2024 | Excel |
| 成长能力_2025 | Excel |
| 总分_斜率 | Excel |
| 总分_趋势 | Excel |
| 盈利能力_斜率 | Excel |
| 盈利能力_趋势 | Excel |
| 成长能力_斜率 | Excel |
| 成长能力_趋势 | Excel |
| 综合趋势 | Excel |

## 子场景(Excel) (1个)

| 字段 | 来源 |
|------|------|
| 定增建议参与 | Excel |

## 标签 (5个)

| 字段 | 来源 |
|------|------|
| 标签_盈利_0 | Excel |
| 标签_盈利_-10 | Excel |
| 标签_盈利_-20 | Excel |
| 7个月涨跌幅 | Excel |
| 最终结论 | Excel |
