# 模型版本: v_20260612_0001_219feat_auc074

**训练时间**: 2026-06-12 00:01
**特征数**: 219
**特征模式**: 全量数值特征
**训练样本**: 1485, 盈利占比: 49.3%
**盈利阈值**: -10.0%

## 训练集指标
| 模型 | 5折CV AUC |
|------|----------|
| LightGBM | 0.741 |
| 逻辑回归 | 0.678 |

## LightGBM 参数
- n_estimators=300, max_depth=5, learning_rate=0.03
- num_leaves=31, is_unbalance=True
- subsample=0.8, colsample_by=0.8, reg_alpha=0.1, reg_lambda=0.1

## 文件清单
- lgb_classifier_full.txt
- lgb_feature_importance_full.csv
- lgb_feature_importance_full.png
- lr_classifier_full.pkl
- lr_coefficients_full.csv
- evaluation_report.txt
- lgb_full_meta.json
