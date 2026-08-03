# 特征工程架构完整分析报告

## 🎯 正确的特征工程架构理解

经过全面分析，现在完全理解了正确的特征工程架构：

### 核心架构

```
ml_features_wide表 (最终特征存储，584个字段)
    ↑
    ├── export_features.py (写入原始特征)
    └── derive_features.py (补充衍生特征)
        ↓
ML训练/回测 (从ml_features_wide表读取)
```

## 📋 各脚本的正确职责

### 1. ml_features_wide表 (核心数据表)
- **字段数量**: 584个字段
- **样本数量**: 568,624个样本
- **作用**: 所有特征的最终存储点
- **包含**: 原始特征 + 衍生特征的统一存储

### 2. export_features.py (原始特征写入)
**职责**: 从MySQL数据源汇总原始特征，写入ml_features_wide表

**数据源**:
- investment_valuation MySQL - 行情/估值/FCF/筛选/行业/定增参数
- placement_evaluation表 - 财务评分/子场景/7个月涨跌幅
- fund_risk_control MySQL - 财务比率/资产负债表/利润表/现金流量表

**输出方式**:
```bash
# 传统方式 (生成parquet文件)
python ml_training/features/export_features.py

# 推荐方式 (直接写入ml_features_wide表)
python ml_training/features/export_features.py --db-register --tag placement_train_20260704
```

**写入字段**: 原始基础特征（股票代码、报价日、PB/PE/PS、财务评分、定增参数等）

### 3. derive_features.py (衍生特征补充)
**职责**: 从基础特征派生高级特征，补充到ml_features_wide表

**特征派生类别**:
- **A类**: FCF增长率 (+18特征)
- **B类**: 动量/波动率 (+12特征)
- **C类**: 行业相对估值 (+8特征)
- **D类**: 市场相对表现 (+6特征)
- **E类**: 时间衰减 (+4特征)
- **F类**: 行业市值占比 (+3特征)
- **G类**: 市场指数相对 (+2特征)

**处理方式**: 从ml_features_wide表读取基础特征 → 计算衍生特征 → 补充回ml_features_wide表

### 4. factor_engine.py (因子计算引擎)
**职责**: 提供标准化因子算子库，支持特征工程

**因子类别**:
- **Beta族**: beta_mkt_{60,120,250}, beta_ind_120, idiovol_120
- **Alpha158缺失族**: K线形态、技术指标、量价、滚动矩

**特点**: 纯函数设计，入参为numpy数组，返回{因子名: 值}

### 5. 其他辅助脚本
- **fetch_factors.py**: 数据获取脚本
- **eval_factors.py**: 因子评估脚本
- **feature_selection.py**: 特征选择脚本
- **feature_exclusions.py**: 特征排除规则

## 🔧 完整的数据流程

### 标准流程
```
1. MySQL数据源
   ↓
2. export_features.py --db-register
   ↓
3. ml_features_wide表 (原始特征)
   ↓
4. derive_features.py
   ↓
5. ml_features_wide表 (原始+衍生特征)
   ↓
6. ML训练/回测分析
```

### 当前问题
1. **PB/PE覆盖率0%**: 需要完成从stock_daily_basic表的回填
2. **流程分散**: 部分脚本仍在生成独立的parquet文件
3. **字段映射**: 需要确认各脚本生成字段与ml_features_wide表字段的对应关系

## 💡 解决方案

### Phase 1: 基础数据完善 (当前阶段)
```bash
# 完成PB/PE回填
python ml_training/scripts/unified_pb_pe_maintenance.py
```

### Phase 2: 标准化export_features.py
```bash
# 使用--db-register参数，直接写入ml_features_wide表
python ml_training/features/export_features.py --db-register --tag placement_train_20260704
```

### Phase 3: 标准化derive_features.py
- 从ml_features_wide表读取基础特征
- 计算衍生特征
- 补充回ml_features_wide表
- 不生成独立的parquet文件

### Phase 4: 清理冗余
```bash
# 删除不需要的中间文件
rm ml_training/data/features.parquet
rm ml_training/data/features_derived.parquet
```

## 📊 关键发现

### 1. ml_features_wide表已经很完整
- **584个字段**: 包含了所有需要的特征
- **568,624个样本**: 涵盖placement和fake_quote两种样本类型
- **字段齐全**: 包含股票代码、报价日、PB/PE/PS、财务评分等所有关键字段

### 2. export_features.py已有DB注册功能
- **--db-register参数**: 可以直接写入ml_features_wide表
- **--tag参数**: 用于标记不同的样本空间
- **--limit参数**: 可以限制处理的样本数量

### 3. 架构设计已经很合理
- **单一数据源**: ml_features_wide表作为所有特征的唯一存储
- **分层处理**: export_features.py负责原始特征，derive_features.py负责衍生特征
- **标准化**: 支持DB注册和parquet文件两种输出方式

## 🎯 下一步行动

### 立即执行
1. **完成PB/PE回填**: 解决0%覆盖率问题
2. **测试DB注册功能**: 验证export_features.py的--db-register功能
3. **确认字段映射**: 确保各脚本生成的字段正确对应ml_features_wide表字段

### 后续优化
1. **标准化流程**: 所有脚本统一使用ml_features_wide表
2. **清理冗余文件**: 删除不需要的中间parquet文件
3. **文档完善**: 更新各脚本的使用说明

---

**总结**: 特征工程架构设计已经很合理，ml_features_wide表是核心，各脚本应该统一使用这个表作为数据存储和读取的中心点。当前的主要任务是解决PB/PE覆盖率问题和标准化各脚本的使用方式。