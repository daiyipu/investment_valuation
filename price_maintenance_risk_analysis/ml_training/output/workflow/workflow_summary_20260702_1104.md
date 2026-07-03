# 模型训练和验证汇总报告

## 基本信息

- **生成时间**: 2026-07-02 11:04:33
- **模型版本**: v_sc_20260701_2111_7m_gray_sc_18feat
- **模型类型**: scorecard
- **期限**: 7m

## 流程完成情况

- ✅ feature_engineering: 完成
- ❌ model_training: 失败
- ❌ model_audit: 失败

## 输出文件位置

### 模型文件
- model_registry.json: 模型元信息注册表
- scorecard_model.pkl: 评分卡模型文件

### 审计报告
- output/audit_<version>/scorecard_audit.csv: 审计数据
- output/audit_<version>/scorecard_audit.md: 可读报告

## 下一步操作

### 如果所有步骤完成 ✅:
1. 查看审计报告确认质量
2. 检查单调性合格率（应>80%）
3. 进行A/B测试对比
4. 生产环境部署

### 如果存在问题 ⚠️:
1. 检查失败步骤的错误信息
2. 使用 `--method tree` 重新审计
3. 使用 `--manual-merge` 优化分箱
4. 重新训练模型

---

*本报告由 train_workflow.py 自动生成*
