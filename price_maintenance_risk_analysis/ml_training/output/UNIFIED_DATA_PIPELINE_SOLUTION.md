# 三套数据流程统一解决方案

## 🎯 问题重新定义

用户提出正确观点：**"用宽表不好吗"**

**核心问题**：不应该生成多个文件，应该统一使用ml_features_wide表作为单一数据源。

## 📊 当前状态分析

### 数据源现状
```
stock_daily_basic表 (原始数据)
├─ 1,687万条记录
├─ 1,676万有PB/PE (99.3%覆盖)
└─ 完整历史数据

ml_features_wide表 (统一数据源)
├─ 568,624个样本
├─ 0% PB/PE覆盖 ❌ 问题所在
└─ 应该是所有流程的唯一数据源

relative_valuation表 (临时数据源)
├─ 1,662条记录
├─ 100% PB/PE覆盖
└─ 不应该单独存在
```

### 三套流程的真实问题

| 流程 | 脚本 | 输出 | 覆盖率 | 问题 |
|------|------|------|--------|------|
| ML训练 | export_features.py | features.parquet | 99.8% | ❌ 不应该生成独立文件 |
| 回测分析 | derive_features.py | features_derived.parquet | 未知 | ❌ 不应该生成独立文件 |
| 数据存储 | backfill_*.py | ml_features_wide表 | 0% | ❌ 数据没有填充 |

## 💡 正确的解决方案

### 统一数据架构

```
                    stock_daily_basic表 (1687万条原始数据)
                               ↓
                        最近交易日匹配逻辑
                               ↓
                    ml_features_wide表 (单一数据源)
                               ↓
        ┌──────────────┴──────────────┐
        ↓                              ↓
    ML训练流程                      回测分析流程
export_features.py              derive_features.py
        ↓                              ↓
  从ml_features_wide读取        从ml_features_wide读取
        ↓                              ↓
    训练模型                        回测验证
```

### 实施步骤

#### Step 1: 修复ml_features_wide表覆盖率 (0% → 85-95%)
```bash
# 执行统一PB/PE回填
python ml_training/scripts/unified_pb_pe_maintenance.py
```

**预期效果**:
- 568,624个样本的PB/PE覆盖率从0%提升到85-95%
- 使用stock_daily_basic表的完整历史数据
- 实施最近交易日匹配逻辑

#### Step 2: 修改export_features.py
```python
# 不要生成独立的features.parquet
# 直接从ml_features_wide表读取数据

def main():
    # 从ml_features_wide表读取数据
    conn = get_db_connection()
    df = pd.read_sql('SELECT * FROM ml_features_wide', conn)
    
    # 直接进行ML训练
    train_model(df)
```

#### Step 3: 修改derive_features.py
```python
# 不要生成独立的features_derived.parquet
# 直接从ml_features_wide表读取数据

def main():
    # 从ml_features_wide表读取数据
    conn = get_db_connection()
    df = pd.read_sql('SELECT * FROM ml_features_wide', conn)
    
    # 直接进行回测分析
    run_backtest(df)
```

#### Step 4: 清理冗余文件
```bash
# 删除不需要的parquet文件
rm ml_training/data/features.parquet
rm ml_training/data/features_derived.parquet  # 如果存在
```

## 🔧 统一维护脚本

### unified_pb_pe_maintenance.py
**功能**: 统一回填ml_features_wide表的PB/PE字段

**核心逻辑**:
1. 从stock_daily_basic表读取完整历史数据
2. 使用最近交易日匹配逻辑
3. 填充到ml_features_wide表的PB/PE字段
4. 实现单一数据源的完整性

### 执行命令
```bash
python ml_training/scripts/unified_pb_pe_maintenance.py
```

### 预期结果
- 568,624个样本的PB/PE覆盖率：0% → 85-95%
- 所有流程统一使用ml_features_wide表
- 不再需要生成多个文件

## 📈 改进效果

### 修改前 (混乱状态)
- ❌ 三个不同的输出文件
- ❌ 数据源不统一
- ❌ 覆盖率不一致
- ❌ 维护复杂

### 修改后 (统一状态)
- ✅ 单一数据源：ml_features_wide表
- ✅ 统一覆盖率：85-95%
- ✅ 标准化流程
- ✅ 简化维护

## 🎯 实施计划

### Phase 1: 数据修复 (立即执行)
1. 执行unified_pb_pe_maintenance.py
2. 验证ml_features_wide表覆盖率提升到85-95%
3. 确认数据质量

### Phase 2: 流程重构 (本周完成)
1. 修改export_features.py，从ml_features_wide表读取
2. 修改derive_features.py，从ml_features_wide表读取
3. 测试ML训练和回测流程

### Phase 3: 清理冗余 (下周完成)
1. 删除不需要的parquet文件
2. 清理废弃脚本
3. 更新文档

## 🔍 关键洞察

用户观点"用宽表不好吗"是正确的：

1. **ml_features_wide表** 568,624个样本，应该成为单一数据源
2. **stock_daily_basic表** 1687万条完整历史数据，应该回填到宽表
3. **不应该生成多个文件**，应该统一使用宽表

这样可以实现：
- 数据源统一
- 覆盖率一致
- 维护简化
- 流程标准化

---

**解决方案状态**: 🟡 执行中
**当前任务**: 正在执行unified_pb_pe_maintenance.py回填ml_features_wide表
**预期效果**: ml_features_wide表覆盖率从0%提升到85-95%
**最终目标**: 建立基于ml_features_wide表的统一数据架构