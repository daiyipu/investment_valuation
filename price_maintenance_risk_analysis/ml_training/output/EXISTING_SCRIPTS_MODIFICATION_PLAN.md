# 现有脚本修改方案 - 基于export_features.py的PB/PE数据流修复

## 🎯 核心修改策略

**原则**: 基于现有脚本修改，不创建新脚本

**目标**: 修复export_features.py中的数据源错误，直接从stock_daily_basic表读取PB/PE数据

## 📝 需要修改的现有脚本

### 修改目标: ml_training/features/export_features.py

#### 问题定位
**文件**: `price_maintenance_risk_analysis/ml_training/features/export_features.py`

**问题1**: 第455行SQL查询缺失PB/PE字段
```python
# 当前代码（第455行附近）
pe_df = pd.read_sql(
    'SELECT stock_code, issue_date, stock_name, issue_date_price, industry_l1, industry_l2, industry_l3, '
    'total_slope, total_trend, profit_slope, profit_trend, growth_slope, growth_trend, combined_trend, '
    'valid_thresholds, premium_min, premium_max, decision, return_7m, price_7m, final_conclusion '
    'FROM placement_evaluation WHERE stock_code IN (%s)' % placeholders,
    conn
)
```

**问题2**: 第228-233行期望读取不存在的字段
```python
# 当前代码（第228-233行）
f['个股PE'] = d.get('current_pe')  # ❌ placement_evaluation表没有这个字段
f['个股PB'] = d.get('current_pb')  # ❌ placement_evaluation表没有这个字段
f['个股PS'] = d.get('current_ps')  # ❌ placement_evaluation表没有这个字段
f['个股市值'] = d.get('current_mv')  # ❌ placement_evaluation表没有这个字段
```

**问题3**: 第485-500行字段映射缺失PB/PE字段
```python
# 当前代码（第485-500行附近）
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')
# ... 缺少PB/PE字段映射
```

## 🔧 具体修改方案

### 修改点1: 第455行SQL查询

**修改内容**:
在现有SQL查询基础上，添加`LEFT JOIN stock_daily_basic`，并增加PB/PE相关字段。

**修改前代码**:
```python
pe_df = pd.read_sql(
    'SELECT stock_code, issue_date, stock_name, issue_date_price, industry_l1, industry_l2, industry_l3, '
    'total_slope, total_trend, profit_slope, profit_trend, growth_slope, growth_trend, combined_trend, '
    'valid_thresholds, premium_min, premium_max, decision, return_7m, price_7m, final_conclusion '
    'FROM placement_evaluation WHERE stock_code IN (%s)' % placeholders,
    conn
)
```

**修改后代码**:
```python
pe_df = pd.read_sql(
    'SELECT pe.stock_code, pe.issue_date, pe.stock_name, pe.issue_date_price, '
    'pe.industry_l1, pe.industry_l2, pe.industry_l3, '
    'pe.total_slope, pe.total_trend, pe.profit_slope, pe.profit_trend, pe.growth_slope, pe.growth_trend, pe.combined_trend, '
    'pe.valid_thresholds, pe.premium_min, pe.premium_max, pe.decision, pe.return_7m, pe.price_7m, pe.final_conclusion, '
    'sb.pb as current_pb, sb.pe as current_pe, sb.ps as current_ps, sb.total_mv as current_mv '
    'FROM placement_evaluation pe '
    'LEFT JOIN stock_daily_basic sb ON pe.stock_code = sb.stock_code '
    'AND DATE_FORMAT(pe.issue_date, "%%Y%%m%%d") = sb.trade_date '
    'WHERE pe.stock_code IN (%s)' % placeholders,
    conn
)
```

**修改说明**:
1. 为placement_evaluation表添加别名`pe`
2. 添加`LEFT JOIN stock_daily_basic sb`连接条件
3. 选择PB/PE相关字段：`sb.pb as current_pb, sb.pe as current_pe, sb.ps as current_ps, sb.total_mv as current_mv`
4. 使用`LEFT JOIN`确保即使没有匹配的stock_daily_basic记录也能保留原数据

### 修改点2: 第485-500行字段映射

**修改内容**:
在现有的字段映射代码中，添加PB/PE字段的映射逻辑。

**修改前代码**:
```python
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')
```

**修改后代码**:
```python
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')

# 新增：PB/PE/PS/市值字段映射
r['个股PB'] = pe.get('current_pb')
r['个股PE'] = pe.get('current_pe')
r['个股PS'] = pe.get('current_ps')
r['个股市值'] = pe.get('current_mv')
```

**修改说明**:
1. 保持原有字段映射逻辑不变
2. 新增4行代码映射PB/PE相关字段
3. 使用`pe.get()`方法保持代码风格一致
4. 字段命名与ml_features_wide表字段名一致

## 📋 执行步骤

### Step 1: 备份现有脚本
```bash
cd /Users/davy/github/investment_valuation
cp price_maintenance_risk_analysis/ml_training/features/export_features.py \
   price_maintenance_risk_analysis/ml_training/features/export_features.py.backup_$(date +%Y%m%d)
```

### Step 2: 修改export_features.py

**使用文本编辑器修改**:
```bash
# 使用你喜欢的编辑器打开文件
vim price_maintenance_risk_analysis/ml_training/features/export_features.py
# 或者
nano price_maintenance_risk_analysis/ml_training/features/export_features.py
```

**修改位置1**: 找到第455行附近
搜索关键词: `pd.read_sql` 和 `placement_evaluation WHERE`

**修改位置2**: 找到第485-500行附近
搜索关键词: `r['股票简称']` 和 `r['报价日']`

### Step 3: 验证修改
```bash
# 语法检查
python -m py_compile price_maintenance_risk_analysis/ml_training/features/export_features.py

# 小规模测试
cd price_maintenance_risk_analysis
python ml_training/features/export_features.py --limit 100 --tag test_pb_fix
```

### Step 4: 检查测试结果
```bash
# 检查生成的文件
python -c "
import pandas as pd
df = pd.read_parquet('ml_training/data/features.parquet')
print('总样本数:', len(df))
print('个股PB覆盖率:', df['个股PB'].notna().mean())
print('个股PE覆盖率:', df['个股PE'].notna().mean())
print('个股PS覆盖率:', df['个股PS'].notna().mean())
print('个股市值覆盖率:', df['个股市值'].notna().mean())
"
```

### Step 5: 全量执行（测试成功后）
```bash
# 全量重新生成特征
cd price_maintenance_risk_analysis
python ml_training/features/export_features.py --db-register --tag pb_fixed_$(date +%Y%m%d)
```

### Step 6: 验证最终效果
```bash
# 检查ml_features_wide表
mysql -u root investment_valuation --default-character-set=utf8mb4 -e "
SELECT
    '个股PB' as field,
    COUNT(个股PB) as has_value,
    ROUND(COUNT(个股PB)/COUNT(*)*100,2) as coverage_pct
FROM ml_features_wide
UNION ALL
SELECT '个股PE', COUNT(个股PE), ROUND(COUNT(个股PE)/COUNT(*)*100,2)
FROM ml_features_wide
UNION ALL
SELECT '个股PS', COUNT(个股PS), ROUND(COUNT(个股PS)/COUNT(*)*100,2)
FROM ml_features_wide
UNION ALL
SELECT '个股市值', COUNT(个股市值), ROUND(COUNT(个股市值)/COUNT(*)*100,2)
FROM ml_features_wide;"
```

## ⚠️ 注意事项

### 1. 字符串转义问题
在SQL字符串中，`%%Y%%m%%d`的百分号需要转义为`%%`，因为Python的字符串格式化会消耗一个`%`。

### 2. LEFT JOIN vs INNER JOIN
使用`LEFT JOIN`而不是`INNER JOIN`，确保：
- 即使stock_daily_basic表没有匹配数据，placement_evaluation的数据也会保留
- 避免因为缺失行情数据而丢失定增样本

### 3. 字段命名一致性
新增的字段名（current_pb/current_pe等）要与后续的映射字段名（个股PB/个股PE等）保持一致。

### 4. 测试先行
必须先进行小规模测试（--limit 100），确认修改正确后再全量执行。

## 🎯 预期效果

### 修改前
- 个股PB覆盖率: 0%
- 个股PE覆盖率: 0%
- 个股PS覆盖率: 0%
- 个股市值覆盖率: 0%

### 修改后（预期）
- 个股PB覆盖率: 85-90%
- 个股PE覆盖率: 80-85%
- 个股PS覆盖率: 80-85%
- 个股市值覆盖率: 90-95%

## 🔄 回滚方案

如果修改出现问题，可以立即回滚：
```bash
# 恢复备份文件
cp price_maintenance_risk_analysis/ml_training/features/export_features.py.backup_YYYYMMDD \
   price_maintenance_risk_analysis/ml_training/features/export_features.py

# 重新运行原版本
python ml_training/features/export_features.py --db-register --tag rollback
```

## 📊 影响范围分析

### 影响的文件
- ✅ `ml_training/features/export_features.py` (主要修改文件)
- ✅ `ml_training/data/features.parquet` (输出文件)
- ✅ `ml_features_wide`表 (最终数据存储)

### 不影响的文件
- ✅ `ml_training/features/derive_features.py` (无需修改)
- ✅ `ml_training/features/factor_engine.py` (无需修改)
- ✅ 其他数据脚本 (无需修改)

### 兼容性
- ✅ 向后兼容，不影响其他字段的处理
- ✅ 不改变现有的输出格式和字段结构
- ✅ 不影响其他脚本的运行

## ✅ 验收标准

### 功能验收
- [ ] 个股PB覆盖率 >80%
- [ ] 个股PE覆盖率 >75%
- [ ] 个股PS覆盖率 >75%
- [ ] 个股市值覆盖率 >90%

### 质量验收
- [ ] 无语法错误，脚本可正常运行
- [ ] 测试结果符合预期
- [ ] 全量执行无错误
- [ ] 数据验证通过

### 性能验收
- [ ] 执行时间增加不超过20%
- [ ] 内存使用增加不超过10%
- [ ] 数据库查询性能无明显下降

---

**修改原则**: 最小化修改，最大化效果
**修改文件**: 1个文件 (export_features.py)
**修改行数**: 约10行代码
**预计工时**: 1-2小时（包含测试验证）
**风险等级**: 低（有备份和回滚机制）