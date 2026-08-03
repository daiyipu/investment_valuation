# 现有脚本全面修改方案 - 解决ml_features_wide表所有字段覆盖问题

## 🎯 问题全景分析

### 发现的所有问题

#### 问题1: PB/PE数据流断裂 ❌
- **现状**: ml_features_wide表PB/PE覆盖率0%，stock_daily_basic表有1687万条完整数据
- **原因**: export_features.py期望从placement_evaluation表读取不存在的current_pe/current_pb字段
- **影响**: 核心估值特征完全缺失

#### 问题2: 其他字段覆盖率低 ⚠️
- **现状**: 584个字段中，约200-300个字段覆盖率<1%
- **原因**: 多种原因：数据源缺失、字段映射错误、数据类型不匹配
- **影响**: 特征工程效果受限，ML训练缺少重要特征

#### 问题3: 数据架构混乱 🔄
- **现状**: placement_evaluation表、company_annual_scores表、stock_daily_basic表数据流混乱
- **原因**: 缺乏统一的数据管理架构
- **影响**: 数据流向不清晰，维护困难

#### 问题4: 数据质量监控缺失 📊
- **现状**: 无法及时发现数据缺失或质量问题
- **原因**: 缺乏数据质量监控机制
- **影响**: 问题发现滞后，影响ML训练效果

## 🔍 完整问题分析

### 高优先级字段覆盖率分析

#### 估值字段组 (优先级：🔴最高)
| 字段名 | 当前覆盖率 | 数据源 | 问题 |
|--------|-----------|--------|------|
| 个股PB | 0% | stock_daily_basic.pb | export_features.py未读取 |
| 个股PE | 0% | stock_daily_basic.pe | export_features.py未读取 |
| 个股PS | 0% | stock_daily_basic.ps | export_features.py未读取 |
| 个股市值 | 0% | stock_daily_basic.total_mv | export_features.py未读取 |
| 同行PB_均值 | 0% | peer_companies表 | 数据连接缺失 |
| 同行PE_均值 | 0% | peer_companies表 | 数据连接缺失 |

#### 财务评分字段组 (优先级：🟡中等)
| 字段名 | 当前覆盖率 | 数据源 | 问题 |
|--------|-----------|--------|------|
| 总分_斜率 | 33% | placement_evaluation表 | 数据部分缺失 |
| 盈利能力_斜率 | 33% | placement_evaluation表 | 数据部分缺失 |
| 成长能力_斜率 | 33% | placement_evaluation表 | 数据部分缺失 |
| 总分_T | 部分覆盖 | company_annual_scores表 | 时间序列映射不完整 |

#### 因子字段组 (优先级：🟢较低)
| 字段名 | 当前覆盖率 | 数据源 | 问题 |
|--------|-----------|--------|------|
| beta_mkt_60 | 低覆盖率 | derive_features.py计算 | 计算逻辑缺失 |
| chip_winner_rate | 低覆盖率 | placement_evaluation表 | 数据摄入不完整 |
| sue_yoy | 低覆盖率 | SUE数据源 | 数据源连接缺失 |

## 📝 现有脚本修改方案

### 修改方案A: export_features.py (主要修改)

#### 修改点1: 第455行SQL查询 - 添加stock_daily_basic连接

**修改前**:
```python
pe_df = pd.read_sql(
    'SELECT stock_code, issue_date, stock_name, issue_date_price, industry_l1, industry_l2, industry_l3, '
    'total_slope, total_trend, profit_slope, profit_trend, growth_slope, growth_trend, combined_trend, '
    'valid_thresholds, premium_min, premium_max, decision, return_7m, price_7m, final_conclusion '
    'FROM placement_evaluation WHERE stock_code IN (%s)' % placeholders,
    conn
)
```

**修改后**:
```python
pe_df = pd.read_sql(
    'SELECT pe.stock_code, pe.issue_date, pe.stock_name, pe.issue_date_price, '
    'pe.industry_l1, pe.industry_l2, pe.industry_l3, '
    'pe.total_slope, pe.total_trend, pe.profit_slope, pe.profit_trend, pe.growth_slope, pe.growth_trend, pe.combined_trend, '
    'pe.valid_thresholds, pe.premium_min, pe.premium_max, pe.decision, pe.return_7m, pe.price_7m, pe.final_conclusion, '
    'sb.pb as current_pb, sb.pe as current_pe, sb.ps as current_ps, sb.total_mv as current_mv, '
    'sw.sw_index_pb, sw.sw_index_pe, sw.sw_index_ps '
    'FROM placement_evaluation pe '
    'LEFT JOIN stock_daily_basic sb ON pe.stock_code = sb.stock_code AND DATE_FORMAT(pe.issue_date, "%%Y%%m%%d") = sb.trade_date '
    'LEFT JOIN sw_index_daily sw ON pe.stock_code = sw.stock_code AND DATE_FORMAT(pe.issue_date, "%%Y%%m%%d") = sw.trade_date '
    'WHERE pe.stock_code IN (%s)' % placeholders,
    conn
)
```

#### 修改点2: 第485-500行字段映射 - 添加估值字段映射

**修改前**:
```python
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')
```

**修改后**:
```python
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')

# 估值字段映射
r['个股PB'] = pe.get('current_pb')
r['个股PE'] = pe.get('current_pe')
r['个股PS'] = pe.get('current_ps')
r['个股市值'] = pe.get('current_mv')
r['行业PB'] = pe.get('sw_index_pb')
r['行业PE'] = pe.get('sw_index_pe')
r['行业PS'] = pe.get('sw_index_ps')
```

### 修改方案B: derive_features.py (次要修改)

#### 检查点: 第310-330行同行数据计算

**现状检查**:
```python
# 检查这段代码是否正确执行
cur.execute('SELECT * FROM peer_companies WHERE stock_code=%s', (code,))
peers = cur.fetchall()
if peers:
    peer_pe = [p['pe'] for p in peers if p.get('pe') is not None]
    peer_pb = [p['pb'] for p in peers if p.get('pb') is not None]
    if peer_pe:
        f['同行PE_均值'] = np.mean(peer_pe)
        f['同行PE_中位'] = np.median(peer_pe)
    if peer_pb:
        f['同行PB_均值'] = np.mean(peer_pb)
        f['同行PB_中位'] = np.median(peer_pb)
```

**可能的问题**:
1. peer_companies表数据是否完整
2. 字段名是否正确（pe/pb vs PE/PB）
3. 数据类型是否匹配

#### 修改点: 增强同行数据计算健壮性

**修改建议**:
```python
# 增加数据验证和错误处理
try:
    cur.execute('SELECT * FROM peer_companies WHERE stock_code=%s', (code,))
    peers = cur.fetchall()
    if peers:
        # 确保字段名正确
        peer_pe = [p.get('pe') or p.get('PE') for p in peers if (p.get('pe') or p.get('PE')) is not None]
        peer_pb = [p.get('pb') or p.get('PB') for p in peers if (p.get('pb') or p.get('PB')) is not None]
        peer_ps = [p.get('ps') or p.get('PS') for p in peers if (p.get('ps') or p.get('PS')) is not None]

        if peer_pe:
            f['同行PE_均值'] = float(np.mean(peer_pe))
            f['同行PE_中位'] = float(np.median(peer_pe))
        if peer_pb:
            f['同行PB_均值'] = float(np.mean(peer_pb))
            f['同行PB_中位'] = float(np.median(peer_pb))
        if peer_ps:
            f['同行PS_均值'] = float(np.mean(peer_ps))
            f['同行PS_中位'] = float(np.median(peer_ps))
except Exception as e:
    print(f"Warning: Failed to calculate peer metrics for {code}: {e}")
```

### 修改方案C: 增强数据验证

#### 在export_features.py末尾添加数据验证函数

```python
def validate_feature_coverage(df, sample_type='all'):
    """验证特征覆盖率"""

    print(f"\n📊 数据覆盖率验证 ({sample_type}):")

    # 高优先级字段
    critical_fields = {
        '个股PB': 90, '个股PE': 85, '个股PS': 85, '个股市值': 95,
        '同行PB_均值': 70, '同行PE_均值': 70, '同行PS_均值': 70,
        '总分_斜率': 80, '盈利能力_斜率': 80, '成长能力_斜率': 80,
    }

    total_samples = len(df)
    coverage_report = {}

    for field, target_coverage in critical_fields.items():
        if field in df.columns:
            actual_coverage = df[field].notna().sum() / total_samples * 100
            status = "✅" if actual_coverage >= target_coverage else "❌"
            print(f"   {status} {field}: {actual_coverage:.1f}% (目标: {target_coverage}%)")
            coverage_report[field] = {
                'actual': actual_coverage,
                'target': target_coverage,
                'status': status
            }
        else:
            print(f"   ⚠️  {field}: 字段不存在")
            coverage_report[field] = {'status': '⚠️ 缺失'}

    return coverage_report
```

#### 在main函数中调用验证

```python
# 在main函数的适当位置添加
if __name__ == '__main__':
    # ... 现有代码 ...

    # 新增：数据覆盖率验证
    if not scored.empty:
        coverage_report = validate_feature_coverage(scored, args.tag or 'all')

        # 生成覆盖率报告文件
        report_path = 'ml_training/output/coverage_report_{}.json'.format(
            args.tag or datetime.now().strftime('%Y%m%d_%H%M%S'))
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(coverage_report, f, indent=2, ensure_ascii=False)
        print(f"   📄 覆盖率报告已保存: {report_path}")
```

## 📋 执行计划

### 阶段1: 诊断和准备 (30分钟)

#### Step 1.1: 字段覆盖率全面检查
```bash
# 检查所有字段的覆盖率
mysql -u root investment_valuation --default-character-set=utf8mb4 -e "
SELECT
    COLUMN_NAME as field_name,
    COUNT(COLUMN_NAME) as has_value,
    ROUND(COUNT(COLUMN_NAME)/COUNT(*)*100,2) as coverage_pct
FROM INFORMATION_SCHEMA.COLUMNS c
JOIN ml_features_wide m ON c.COLUMN_NAME = m.COLUMN_NAME
WHERE c.TABLE_SCHEMA = 'investment_valuation'
AND c.TABLE_NAME = 'ml_features_wide'
GROUP BY COLUMN_NAME
ORDER BY coverage_pct ASC
LIMIT 50;"
```

#### Step 1.2: 数据源映射分析
```bash
# 分析各字段的数据源
python -c "
# 创建数据源映射分析脚本
import pymysql
conn = pymysql.connect(host='127.0.0.1', user='root', password='',
                      database='investment_valuation', charset='utf8mb4')

# 检查各表的数据情况
tables = ['placement_evaluation', 'company_annual_scores', 'stock_daily_basic',
          'peer_companies', 'sw_index_daily', 'industry_daily']

for table in tables:
    cur = conn.cursor()
    cur.execute(f'SELECT COUNT(*) FROM {table}')
    count = cur.fetchone()[0]
    print(f'{table}: {count:,}条记录')
    cur.close()
"
```

### 阶段2: 核心修改 (2小时)

#### Step 2.1: 备份现有脚本
```bash
cd /Users/davy/github/investment_valuation
cp price_maintenance_risk_analysis/ml_training/features/export_features.py \
   price_maintenance_risk_analysis/ml_training/features/export_features.py.backup_$(date +%Y%m%d_%H%M%S)

cp price_maintenance_risk_analysis/ml_training/features/derive_features.py \
   price_maintenance_risk_analysis/ml_training/features/derive_features.py.backup_$(date +%Y%m%d_%H%M%S)
```

#### Step 2.2: 修改export_features.py
1. 修改第455行SQL查询（添加stock_daily_basic和sw_index_daily连接）
2. 修改第485-500行字段映射（添加估值字段）
3. 添加数据验证函数

#### Step 2.3: 修改derive_features.py
1. 检查第310-330行同行数据计算逻辑
2. 增强数据验证和错误处理

### 阶段3: 测试验证 (1小时)

#### Step 3.1: 小规模测试
```bash
cd price_maintenance_risk_analysis
python ml_training/features/export_features.py --limit 100 --tag test_all_fields
```

#### Step 3.2: 检查测试结果
```bash
# 检查生成的特征文件
python -c "
import pandas as pd
import json

df = pd.read_parquet('ml_training/data/features.parquet')
print('总样本数:', len(df))
print('字段数:', len(df.columns))

# 检查关键字段覆盖率
critical_fields = ['个股PB', '个股PE', '个股PS', '个股市值',
                  '同行PB_均值', '同行PE_均值', '总分_斜率']

for field in critical_fields:
    if field in df.columns:
        coverage = df[field].notna().mean() * 100
        print(f'{field}: {coverage:.1f}%')
    else:
        print(f'{field}: 字段缺失')

# 读取覆盖率报告
try:
    with open('ml_training/output/coverage_report_*.json', 'r') as f:
        report = json.load(f)
        print('\\n覆盖率报告:', json.dumps(report, indent=2, ensure_ascii=False))
except:
    print('未找到覆盖率报告')
"
```

#### Step 3.3: 全量测试（小规模测试通过后）
```bash
cd price_maintenance_risk_analysis
python ml_training/features/export_features.py --db-register --tag all_fields_fixed_$(date +%Y%m%d)
```

### 阶段4: 效果验证 (30分钟)

#### Step 4.1: 验证ml_features_wide表
```bash
mysql -u root investment_valuation --default-character-set=utf8mb4 -e "
SELECT '总体效果' as metric,
       COUNT(*) as total_samples,
       COUNT(个股PB) as has_pb,
       ROUND(COUNT(个股PB)/COUNT(*)*100,2) as pb_coverage,
       COUNT(个股PE) as has_pe,
       ROUND(COUNT(个股PE)/COUNT(*)*100,2) as pe_coverage
FROM ml_features_wide

UNION ALL

SELECT '高优先级字段效果',
       COUNT(*) as total_samples,
       SUM(CASE WHEN 个股PB IS NOT NULL THEN 1 ELSE 0 END +
           CASE WHEN 个股PE IS NOT NULL THEN 1 ELSE 0 END +
           CASE WHEN 个股PS IS NOT NULL THEN 1 ELSE 0 END +
           CASE WHEN 个股市值 IS NOT NULL THEN 1 ELSE 0 END) as has_pb,
       0 as pb_coverage,
       0 as has_pe,
       0 as pe_coverage
FROM ml_features_wide;"
```

#### Step 4.2: 对比修改前后效果
```bash
# 生成对比报告
python -c "
print('📊 修改效果对比报告')
print('=' * 50)
print('修改前 -> 修改后')
print('个股PB覆盖率: 0% -> 85%+ (预期)')
print('个股PE覆盖率: 0% -> 80%+ (预期)')
print('同行PB_均值覆盖率: 0% -> 70%+ (预期)')
print('总分_斜率覆盖率: 33% -> 80%+ (预期)')
"
```

## 🎯 预期效果

### 核心字段覆盖率提升
| 字段组 | 修改前 | 修改后(预期) | 提升 |
|--------|-------|------------|------|
| 个股估值字段(PB/PE/PS/市值) | 0% | 85%+ | +85% |
| 同行估值字段 | 0% | 70%+ | +70% |
| 财务评分字段 | 33% | 80%+ | +47% |
| 行业估值字段 | 90% | 95%+ | +5% |

### 整体数据质量提升
- **字段完整度**: 从40%完整 → 90%完整
- **数据准确性**: 直接从源表获取，准确率100%
- **数据一致性**: 统一数据口径，消除不一致
- **可维护性**: 建立标准化数据流程

## ⚠️ 风险控制

### 技术风险
- **数据源依赖**: 确保stock_daily_basic等源表数据完整
- **性能影响**: LEFT JOIN可能增加查询时间，需监控
- **数据类型**: 确保字段类型匹配，避免类型转换错误

### 回滚机制
```bash
# 如遇问题，立即回滚
cp price_maintenance_risk_analysis/ml_training/features/export_features.py.backup_YYYYMMDD_HHMMSS \
   price_maintenance_risk_analysis/ml_training/features/export_features.py

# 重新运行原版本
python ml_training/features/export_features.py --db-register --tag rollback
```

### 分步执行
1. 先修改export_features.py，测试验证
2. 再修改derive_features.py，测试验证
3. 每步修改后都要验证效果
4. 遇到问题立即回滚，分析原因

## ✅ 验收标准

### 功能验收
- [ ] 个股PB覆盖率 >80%
- [ ] 个股PE覆盖率 >75%
- [ ] 同行估值字段覆盖率 >70%
- [ ] 财务评分字段覆盖率 >75%

### 质量验收
- [ ] 脚本无语法错误，可正常运行
- [ ] 测试结果符合预期
- [ ] 数据验证无异常
- [ ] 覆盖率报告生成正常

### 业务验收
- [ ] ML训练可使用完整特征
- [ ] 模型性能不下降
- [ ] 数据流程清晰可维护
- [ ] 问题可快速定位

---

**修改原则**: 基于现有脚本，最小化修改，最大化效果
**修改文件**: 2个文件 (export_features.py, derive_features.py)
**修改范围**: 约20行代码
**预计工时**: 4小时 (诊断30min + 修改2h + 测试1h + 验证30min)
**风险等级**: 中等 (有备份和回滚机制，分步执行验证)
**预期收益**: 核心字段覆盖率从当前水平提升到75%+