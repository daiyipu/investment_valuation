# v1 — 320特征全量模型

**训练时间**: 2026-06-11
**模型文件**: 全量数值特征 (320个)

## 训练配置
- 模型: LightGBM (is_unbalance=True, n_estimators=300, max_depth=5, lr=0.03, num_leaves=31)
- 训练样本: 1236条, 盈利占比 43.9%
- 阈值: 标签_盈利_-10 (涨跌幅 > -10%)
- 特征: features_derived.parquet 全部数值列 (排除标签)

## 训练集指标
| 模型 | 5折CV AUC |
|------|----------|
| LightGBM | 0.746 ± 0.034 |
| 逻辑回归 | 0.655 ± 0.023 |

## 测试集指标 (询转项目, 399条)
| 指标 | LightGBM | 逻辑回归 |
|------|----------|---------|
| AUC (盈利>0) | 0.675 | 0.661 |
| AUC (盈利>-10%) | 0.687 | 0.654 |
| KS (盈利>-10%) | 0.426 | 0.265 |
| 精确率 (盈利>-10%) | 93.6% | 90.6% |
| 召回率 (盈利>-10%) | 57.9% | 47.0% |
| 概率vs浮盈相关系数 | 0.224 | — |

## 特征重要性 Top 10
1. 市场波动率_120d (107) — G类大盘
2. 报价日_excel (102)
3. 市场波动率比值 (88) — G类大盘
4. 波动率_60d (71)
5. 报价日_md (70)
6. price_vs_MA20 (68) — E类行情动量
7. 行业胜率_120d (58)
8. 行业年化收益_60d (56)
9. 盈利能力_delta_1y (52) — C类评分变动
10. 同行市值_均值 (51)

## 文件清单
- lgb_classifier_full.txt — LightGBM模型
- lgb_full_meta.json — 特征列表 + median (predict_profitability使用)
- lr_classifier_full.pkl — 逻辑回归模型 (含scaler, features, medians)
- lr_coefficients_full.csv — 逻辑回归系数表(评分卡)
- lgb_feature_importance_full.csv — 特征重要性
- lgb_feature_importance_full.png — 特征重要性图
- evaluation_report.txt — 评估报告
- test_validation_result.csv — 测试集验证结果
