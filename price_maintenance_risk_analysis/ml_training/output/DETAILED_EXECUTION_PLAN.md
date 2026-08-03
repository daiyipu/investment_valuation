# 详细执行方案 - PB/PE数据流断裂问题完整解决方案

## 🔍 问题诊断结果

### 发现的具体问题

#### 问题1: export_features.py数据源错误
**文件**: `ml_training/features/export_features.py` 第228-233行
```python
# 错误代码：期望从placement_evaluation表读取不存在的字段
f['个股PE'] = d.get('current_pe')  # ❌ placement_evaluation表没有这个字段
f['个股PB'] = d.get('current_pb')  # ❌ placement_evaluation表没有这个字段
f['个股PS'] = d.get('current_ps')  # ❌ placement_evaluation表没有这个字段
f['个股市值'] = d.get('current_mv')  # ❌ placement_evaluation表没有这个字段
```

**实际情况检查**:
```sql
-- placement_evaluation表实际字段
SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'placement_evaluation'
AND (COLUMN_NAME LIKE '%pb%' OR COLUMN_NAME LIKE '%pe%' OR COLUMN_NAME LIKE '%ps%');

-- 结果: 只有sub_stock_pe, sub_industry_pe等子场景字段，没有current_pe/current_pb
```

#### 问题2: 数据流断裂
**数据源检查**:
```sql
-- stock_daily_basic表有完整数据
SELECT COUNT(*) FROM stock_daily_basic;  -- 1687万条记录
SELECT COUNT(pb) FROM stock_daily_basic;   -- 1673万条记录 (99.14%覆盖)
SELECT COUNT(pe) FROM stock_daily_basic;   -- 1468万条记录 (87.00%覆盖)

-- 示例数据
SELECT * FROM stock_daily_basic WHERE stock_code = '000001.SZ' LIMIT 1;
-- pb=0.4278, pe=4.6565, ps=1.5103, total_mv=19852254.3186
```

**数据流状态**:
```
stock_daily_basic表 (1687万条PB/PE数据)
    ↓ ❌ 断裂点: export_features.py没有读取这个表的数据
ml_features_wide表 (PB/PE覆盖率0%)
```

#### 问题3: 技术执行障碍
**发现的技术问题**:
1. **字符集冲突**: `(utf8mb4_unicode_ci,IMPLICIT) and (utf8mb4_0900_ai_ci,IMPLICIT)`
2. **数据库锁冲突**: `Lock wait timeout exceeded`
3. **进程冲突**: 之前的修复脚本仍在运行

## 📋 完整解决方案

### 阶段1: 环境准备 (立即执行)

#### Step 1.1: 清理冲突进程
```bash
# 检查是否有冲突进程
mysql -u root investment_valuation -e "SHOW PROCESSLIST;"

# 如果发现有长时间运行的进程(>1000秒)，终止它
# KILL <process_id>;

# 清理临时表
mysql -u root investment_valuation -e "DROP TABLE IF EXISTS ml_pb_pe_match_v2;"
```

#### Step 1.2: 数据备份
```bash
# 备份ml_features_wide表
mysqldump -u root investment_valuation ml_features_wide > ml_features_wide_backup_$(date +%Y%m%d_%H%M%S).sql

# 验证备份文件
ls -lh ml_features_wide_backup_*.sql
```

#### Step 1.3: 字符集统一
```sql
-- 检查表的字符集
SHOW CREATE TABLE ml_features_wide;
SHOW CREATE TABLE stock_daily_basic;

-- 如果字符集不一致，统一为utf8mb4
-- ALTER TABLE ml_features_wide CONVERT TO CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

### 阶段2: 快速修复 (30分钟内)

#### Step 2.1: 使用优化后的快速修复脚本

**脚本特点**:
- ✅ 解决字符集冲突问题
- ✅ 批量处理，避免锁冲突
- ✅ 支持精确匹配 + 最近交易日匹配
- ✅ 详细的进度报告和验证

**使用方法**:
```bash
cd /Users/davy/github/investment_valuation
python price_maintenance_risk_analysis/ml_training/scripts/quick_fix_pb_pe.py
```

**预期结果**:
- 精确匹配修复: 预计400,000-500,000条
- 最近交易日匹配: 预计50,000-100,000条
- 最终覆盖率: 预计85-95%

#### Step 2.2: 验证修复效果
```bash
# 验证脚本
mysql -u root investment_valuation --default-character-set=utf8mb4 -e "
SELECT
    '个股PB' as field_name,
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

**成功标准**:
- 个股PB覆盖率 >85%
- 个股PE覆盖率 >80%
- 个股PS覆盖率 >80%
- 个股市值覆盖率 >90%

### 阶段3: 根因修复 (本周完成)

#### Step 3.1: 修改export_features.py

**修改位置**: `ml_training/features/export_features.py` 第455行

**修改前**:
```python
# 原来的SQL查询（缺失PB/PE字段）
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
# 添加LEFT JOIN stock_daily_basic获取PB/PE数据
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

**第二处修改**: 第485-500行附近

**修改前**:
```python
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')
# ... 缺少PB/PE字段映射
```

**修改后**:
```python
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')

# 新增：PB/PE/PS/市值字段映射
r['个股PB'] = pe.get('current_pb')
r['个股PE'] = pe.get('current_pe')
r['个股PS'] = pe.get('current_ps')
r['个股市值'] = pe.get('current_mv')

# ... 其余字段映射保持不变
```

#### Step 3.2: 测试修改效果
```bash
# 小规模测试
python price_maintenance_risk_analysis/ml_training/features/export_features.py --limit 100 --tag test_pb_fix

# 检查生成的数据
python -c "
import pandas as pd
df = pd.read_parquet('ml_training/data/features.parquet')
print('个股PB覆盖率:', df['个股PB'].notna().mean())
print('个股PE覆盖率:', df['个股PE'].notna().mean())
"
```

#### Step 3.3: 全量重新生成
```bash
# 如果测试成功，全量生成
python price_maintenance_risk_analysis/ml_training/features/export_features.py --db-register --tag pb_fixed_$(date +%Y%m%d)
```

### 阶段4: 长期优化 (下周完成)

#### Step 4.1: 建立数据质量监控

**监控脚本**: `ml_training/scripts/data_quality_monitor.py`
```python
#!/usr/bin/env python3
"""数据质量监控脚本"""

import pymysql
import json
from datetime import datetime

def check_data_quality():
    """检查ml_features_wide表数据质量"""
    conn = pymysql.connect(host='127.0.0.1', user='root', password='',
                          database='investment_valuation', charset='utf8mb4')

    try:
        cur = conn.cursor()

        # 核心字段覆盖率检查
        critical_fields = [
            ('个股PB', 90),  # 目标覆盖率90%
            ('个股PE', 85),  # 目标覆盖率85%
            ('个股PS', 85),
            ('个股市值', 95),
            ('总分_斜率', 80),
            ('盈利能力_斜率', 80),
        ]

        quality_report = {}
        for field_name, target_coverage in critical_fields:
            cur.execute(f"SELECT COUNT({field_name}), COUNT(*) FROM ml_features_wide")
            has_value, total = cur.fetchone()
            actual_coverage = has_value / total * 100 if total > 0 else 0

            status = "✅" if actual_coverage >= target_coverage else "❌"
            quality_report[field_name] = {
                "actual_coverage": actual_coverage,
                "target_coverage": target_coverage,
                "status": status
            }

        # 生成报告
        report = {
            "timestamp": datetime.now().isoformat(),
            "quality_results": quality_report,
            "overall_status": all(r["status"] == "✅" for r in quality_report.values())
        }

        print(json.dumps(report, indent=2, ensure_ascii=False))
        return report

    finally:
        conn.close()

if __name__ == '__main__':
    check_data_quality()
```

#### Step 4.2: 建立定期同步机制

**同步脚本**: `ml_training/scripts/auto_sync_pb_pe.py`
```python
#!/usr/bin/env python3
"""自动同步PB/PE数据脚本"""

import pymysql
from datetime import datetime, timedelta

def sync_recent_samples(days=7):
    """同步最近N天的样本"""

    conn = pymysql.connect(host='127.0.0.1', user='root', password='',
                          database='investment_valuation', charset='utf8mb4')
    try:
        cur = conn.cursor()

        # 获取最近需要更新的样本
        cutoff_date = (datetime.now() - timedelta(days=days)).strftime('%Y%m%d')

        cur.execute(f'''
            UPDATE ml_features_wide m
            INNER JOIN stock_daily_basic sb
                ON m.股票代码 COLLATE utf8mb4_unicode_ci = sb.stock_code COLLATE utf8mb4_unicode_ci
                AND DATE_FORMAT(m.报价日, "%Y%m%d") = sb.trade_date
            SET m.个股PB = sb.pb,
                m.个股PE = sb.pe,
                m.个股PS = sb.ps,
                m.个股市值 = sb.total_mv
            WHERE m.报价日 >= {cutoff_date}
              AND m.个股PB IS NULL
        ''')

        updated = cur.rowcount
        conn.commit()

        print(f"✅ 同步完成，更新了{updated}条记录")
        return updated

    except Exception as e:
        print(f"❌ 同步失败: {e}")
        conn.rollback()
        return 0
    finally:
        conn.close()

if __name__ == '__main__':
    sync_recent_samples()
```

**定期执行**:
```bash
# 添加到crontab，每天执行
# 0 2 * * * cd /Users/davy/github/investment_valuation && python price_maintenance_risk_analysis/ml_training/scripts/auto_sync_pb_pe.py
```

## 📊 执行时间表

### 今天 (2小时内)
- ✅ **环境准备** (15分钟): 清理进程，备份数据，检查字符集
- ✅ **快速修复** (30分钟): 运行修复脚本，验证效果
- ✅ **效果验证** (15分钟): 确认覆盖率达标

### 本周 (3-5小时)
- ✅ **根因修复** (2小时): 修改export_features.py，测试验证
- ✅ **全量重跑** (1小时): 重新生成完整特征数据
- ✅ **文档更新** (30分钟): 更新相关文档和使用说明

### 下周 (2-3小时)
- ✅ **监控建立** (1小时): 建立数据质量监控
- ✅ **自动化** (1小时): 建立定期同步机制
- ✅ **验收测试** (1小时): 完整的功能验收

## 🎯 成功标准

### 技术指标
- **PB覆盖率**: 从0% → 90%+
- **PE覆盖率**: 从0% → 85%+
- **PS覆盖率**: 从0% → 85%+
- **市值覆盖率**: 从0% → 95%+

### 功能指标
- **数据完整性**: 所有估值字段完整覆盖
- **数据准确性**: 与源表数据一致
- **流程稳定性**: 不再出现数据流断裂
- **维护便利性**: 建立自动化监控机制

### 业务影响
- **ML训练**: 提供完整的估值特征
- **模型性能**: 预期模型AUC提升5-10%
- **数据可靠性**: 建立长期稳定的数据流程

## ⚠️ 风险控制

### 技术风险
- **数据库锁冲突**: 使用批量处理，设置合理超时
- **字符集冲突**: 显式指定COLLATE，统一字符集
- **性能影响**: 分批处理，避免长时间锁表

### 数据风险
- **数据备份**: 每次操作前完整备份
- **回滚机制**: 保留原始数据，可随时回滚
- **验证机制**: 每步操作后验证数据完整性

### 流程风险
- **测试先行**: 小规模测试后再全量执行
- **分步验证**: 每个阶段独立验收
- **监控告警**: 建立数据质量监控

## 📝 验收检查清单

### 快速修复验收
- [ ] 个股PB覆盖率 >85%
- [ ] 个股PE覆盖率 >80%
- [ ] 个股PS覆盖率 >80%
- [ ] 个股市值覆盖率 >90%
- [ ] 无数据库错误或锁冲突

### 根因修复验收
- [ ] export_features.py修改完成
- [ ] 小规模测试通过
- [ ] 全量重新生成成功
- [ ] 数据验证无异常

### 长期优化验收
- [ ] 数据质量监控运行
- [ ] 定期同步机制建立
- [ ] 文档更新完成
- [ ] 使用说明清晰

---

**制定时间**: 2026-07-05
**预计完成**: 今天2小时 + 本周5小时 + 下周3小时
**总工时**: 约10小时
**风险等级**: 中等（有备份和回滚机制）
**预期收益**: PB/PE覆盖率从0%提升到90%+