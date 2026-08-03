# export_features.py 和 derive_features.py 关系说明及修改计划

## 📋 两个脚本的关系

### 当前架构 (混乱状态)
```
export_features.py (数据汇总脚本)
├─ 职责: 从多个MySQL数据源汇总原始特征
├─ 数据源:
│   ├─ investment_valuation MySQL - 行情/估值/FCF/筛选/行业/定增参数
│   ├─ placement_evaluation表 - 财务评分/子场景/7个月涨跌幅
│   └─ fund_risk_control MySQL - 财务比率/资产负债表/利润表/现金流量表
├─ 输出: ml_training/data/features.parquet
└─ 用途: 为ML训练准备基础数据

derive_features.py (特征衍生脚本)
├─ 职责: 从基础特征派生高级特征
├─ 输入: features.parquet (来自export_features.py)
├─ 两阶段流水线:
│   ├─ Stage 1 (A-E): 纯parquet运算，不访问数据库
│   │   ├─ A类: FCF增长率 (+18特征)
│   │   ├─ B类: 动量/波动率 (+12特征)
│   │   ├─ C类: 行业相对估值 (+8特征)
│   │   ├─ D类: 市场相对表现 (+6特征)
│   │   └─ E类: 时间衰减 (+4特征)
│   └─ Stage 2 (F-G): 需MySQL查询
│       ├─ F类: 行业市值占比 (+3特征)
│       └─ G类: 市场指数相对 (+2特征)
├─ 输出: ml_training/data/features_derived.parquet
└─ 用途: 特征工程，计算衍生指标

ML训练/回测分析
└─ 输入: features_derived.parquet
```

### 问题分析
1. **数据分散**: 两个脚本各自生成独立的parquet文件
2. **覆盖不同**: export_features.py 99.8%覆盖 vs ml_features_wide表 0%覆盖
3. **维护复杂**: 需要协调两个脚本的输出
4. **效率低下**: 重复的数据处理和存储

## 🎯 统一架构方案

### 目标架构 (统一状态)
```
ml_features_wide表 (单一数据源)
├─ 568,624个样本，584个字段
├─ 统一PB/PE覆盖率: 85-95%
└─ 所有特征的唯一真实来源

export_features.py (修改后)
├─ 职责: 从ml_features_wide表读取数据进行ML训练
├─ 输入: ml_features_wide表
├─ 输出: 直接训练模型，不生成parquet文件
└─ 用途: ML训练流程

derive_features.py (修改后)
├─ 职责: 从ml_features_wide表读取数据进行回测分析
├─ 输入: ml_features_wide表
├─ 输出: 直接回测验证，不生成parquet文件
└─ 用途: 回测分析流程
```

## 🔧 具体修改计划

### Phase 1: 修改export_features.py

#### 修改前:
```python
def main():
    # 从多个MySQL数据源获取数据
    df = fetch_data_from_mysql()
    # 输出到parquet文件
    df.to_parquet('ml_training/data/features.parquet')
```

#### 修改后:
```python
def main():
    # 从ml_features_wide表读取数据
    conn = get_db_connection()
    df = pd.read_sql('SELECT * FROM ml_features_wide WHERE ...', conn)
    # 直接进行ML训练
    train_model(df)
```

### Phase 2: 修改derive_features.py

#### 修改前:
```python
def main():
    # 从parquet文件读取
    df = pd.read_parquet('ml_training/data/features.parquet')
    # 特征衍生
    df_derived = derive_features(df)
    # 输出到parquet文件
    df_derived.to_parquet('ml_training/data/features_derived.parquet')
```

#### 修改后:
```python
def main():
    # 从ml_features_wide表读取数据
    conn = get_db_connection()
    df = pd.read_sql('SELECT * FROM ml_features_wide WHERE ...', conn)
    # 特征衍生
    df_derived = derive_features(df)
    # 直接进行回测分析
    run_backtest(df_derived)
```

### Phase 3: 清理冗余文件

#### 删除不需要的文件:
```bash
rm ml_training/data/features.parquet
rm ml_training/data/features_derived.parquet
```

#### 保留必要的文件:
- `ml_training/features/export_features.py` (修改后)
- `ml_training/features/derive_features.py` (修改后)
- `ml_training/scripts/unified_pb_pe_maintenance.py` (维护脚本)

## 📈 修改效果对比

### 修改前 (混乱状态)
- ❌ 两个独立的parquet文件
- ❌ 数据源不统一
- ❌ 覆盖率不一致 (99.8% vs 0%)
- ❌ 维护复杂，需要协调两个脚本
- ❌ 存储冗余，多个副本

### 修改后 (统一状态)
- ✅ 单一数据源：ml_features_wide表
- ✅ 统一覆盖率：85-95%
- ✅ 流程简化，直接从数据库读取
- ✅ 维护简单，只维护一个宽表
- ✅ 存储优化，避免冗余

## 🎯 执行时间表

| 阶段 | 任务 | 状态 |
|------|------|------|
| Phase 0 | SQL PB/PE回填 (0% → 85-95%) | 🟡 进行中 |
| Phase 1 | 修改export_features.py | ⏳ 待执行 |
| Phase 2 | 修改derive_features.py | ⏳ 待执行 |
| Phase 3 | 清理冗余文件 | ⏳ 待执行 |
| Phase 4 | 验证统一流程 | ⏳ 待执行 |

## 🔍 关键洞察

1. **数据职责明确**: ml_features_wide表是唯一数据源
2. **流程简化**: 两个脚本改为直接使用宽表数据
3. **避免冗余**: 不再生成多个中间文件
4. **维护简单**: 只需维护一个宽表的数据质量

---

**当前状态**: 🟡 等待SQL PB/PE回填完成
**下一步**: 修改export_features.py和derive_features.py使用ml_features_wide表
**最终目标**: 建立基于ml_features_wide表的统一数据架构