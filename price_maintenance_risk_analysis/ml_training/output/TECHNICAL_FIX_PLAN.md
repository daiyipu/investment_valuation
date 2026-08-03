# 具体技术修复方案 - 解决PB/PE数据流断裂问题

## 🔍 问题根因分析

### 发现的具体断裂点

#### 断裂点1: export_features.py期望字段不存在
**文件**: `ml_training/features/export_features.py` 第228-233行
```python
# 代码期望从placement_evaluation表读取：
f['个股PE'] = d.get('current_pe')  # ❌ placement_evaluation表没有这个字段！
f['个股PB'] = d.get('current_pb')  # ❌ placement_evaluation表没有这个字段！
f['个股PS'] = d.get('current_ps')  # ❌ placement_evaluation表没有这个字段！
f['个股市值'] = d.get('current_mv')  # ❌ placement_evaluation表没有这个字段！
```

**实际情况**:
- placement_evaluation表没有current_pe/current_pb/current_ps/current_mv字段
- 只有sub_stock_pe/sub_industry_pe等子场景字段
- 导致所有估值字段返回None，最终ml_features_wide表覆盖率为0%

#### 断裂点2: stock_daily_basic数据未被利用
**数据现状**:
```sql
-- stock_daily_basic表有完整数据（1687万条记录）
SELECT pb, pe, ps, total_mv FROM stock_daily_basic WHERE stock_code = '000001.SZ';
-- 结果: pb=0.4278, pe=4.6565, ps=1.5103, total_mv=19852254.3186
```

**数据丢失**:
- stock_daily_basic表有完整且准确的PB/PE/PS/市值数据
- 但export_features.py没有从这个表读取这些字段
- 数据流断裂: stock_daily_basic → ❌ → ml_features_wide

## 💡 具体修复方案

### 方案A: 修改export_features.py数据源（推荐）

#### 修复策略
直接从stock_daily_basic表读取PB/PE数据，绕过placement_evaluation表

#### 具体代码修改

**位置**: `ml_training/features/export_features.py` 第455行附近

**修改前**:
```python
# 当前的错误代码（第455行）
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
# 修复后的代码
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

**第二处修改**: `ml_training/features/export_features.py` 第485-500行附近

**修改前**:
```python
# 原来的字段映射逻辑（缺失PB/PE字段）
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')
# ... 其他字段映射，但没有PB/PE/PS/市值字段
```

**修改后**:
```python
# 修复后的字段映射逻辑
r['股票简称'] = pe.get('stock_name', '')
r['报价日'] = pe.get('issue_date')
r['报价日价格'] = pe.get('issue_date_price')

# 新增：PB/PE/PS/市值字段映射
r['个股PB'] = pe.get('current_pb')
r['个股PE'] = pe.get('current_pe')
r['个股PS'] = pe.get('current_ps')
r['个股市值'] = pe.get('current_mv')

# ... 其他字段映射保持不变
```

#### 修改理由
1. **最小改动原则**: 只修改SQL查询，不改变其他逻辑
2. **数据准确性**: 直接从源表获取数据，避免中间环节错误
3. **向后兼容**: 不影响其他字段的读取逻辑
4. **性能考虑**: 使用LEFT JOIN，即使stock_daily_basic没有匹配也能继续

### 方案B: 直接修复ml_features_wide表数据（快速修复）

#### 修复策略
绕过export_features.py，直接回填ml_features_wide表的PB/PE字段

#### 具体修复脚本
```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速修复ml_features_wide表的PB/PE/PS/市值字段
直接从stock_daily_basic表回填数据
"""

import pymysql
import pandas as pd
import numpy as np
from tqdm import tqdm

def quick_fix_pb_pe_data():
    """快速修复PB/PE数据"""
    
    print("🚀 开始快速修复PB/PE数据...")
    
    conn = pymysql.connect(host='127.0.0.1', user='root', password='',
                          database='investment_valuation', charset='utf8mb4')
    
    try:
        # Step 1: 检查当前数据状态
        print("\n📊 检查当前数据状态...")
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM ml_features_wide")
        total_samples = cur.fetchone()[0]
        
        cur.execute("SELECT COUNT(个股PB) FROM ml_features_wide WHERE 个股PB IS NOT NULL")
        current_coverage = cur.fetchone()[0]
        
        print(f"   总样本数: {total_samples:,}")
        print(f"   当前PB覆盖率: {current_coverage:,} ({current_coverage/total_samples*100:.2f}%)")
        
        # Step 2: 精确匹配修复（报价日 = 交易日）
        print("\n🔧 Step 1: 精确匹配修复...")
        cur.execute('''
            UPDATE ml_features_wide m
            INNER JOIN stock_daily_basic sb
                ON m.股票代码 = sb.stock_code
                AND DATE_FORMAT(m.报价日, "%%Y%%m%%d") = sb.trade_date
            SET m.个股PB = sb.pb,
                m.个股PE = sb.pe,
                m.个股PS = sb.ps,
                m.个股市值 = sb.total_mv
            WHERE m.个股PB IS NULL
        ''')
        exact_match = cur.rowcount
        conn.commit()
        print(f"   精确匹配修复: {exact_match:,}条")
        
        # Step 3: 最近交易日匹配（报价日 != 交易日）
        print("\n🔧 Step 2: 最近交易日匹配...")
        for days_back in range(1, 6):
            cur.execute(f'''
                UPDATE ml_features_wide m
                INNER JOIN stock_daily_basic sb
                    ON m.股票代码 = sb.stock_code
                    AND DATE_FORMAT(DATE_SUB(m.报价日, INTERVAL {days_back} DAY), "%%Y%%m%%d") = sb.trade_date
                SET m.个股PB = sb.pb,
                    m.个股PE = sb.pe,
                    m.个股PS = sb.ps,
                    m.个股市值 = sb.total_mv
                WHERE m.个股PB IS NULL
            ''')
            nearest_match = cur.rowcount
            conn.commit()
            print(f"   Day -{days_back}: {nearest_match:,}条")
            
            if nearest_match == 0:
                break
        
        # Step 4: 验证修复效果
        print("\n📊 验证修复效果...")
        cur.execute("SELECT COUNT(个股PB) FROM ml_features_wide WHERE 个股PB IS NOT NULL")
        final_coverage = cur.fetchone()[0]
        coverage_pct = final_coverage / total_samples * 100
        
        print(f"   最终PB覆盖率: {final_coverage:,} ({coverage_pct:.2f}%)")
        
        # 详细统计
        cur.execute('''
            SELECT 
                COUNT(个股PB) as has_pb,
                COUNT(个股PE) as has_pe, 
                COUNT(个股PS) as has_ps,
                COUNT(个股市值) as has_mv
            FROM ml_features_wide
        ''')
        stats = cur.fetchone()
        print(f"   个股PB: {stats[0]:,} ({stats[0]/total_samples*100:.2f}%)")
        print(f"   个股PE: {stats[1]:,} ({stats[1]/total_samples*100:.2f}%)")
        print(f"   个股PS: {stats[2]:,} ({stats[2]/total_samples*100:.2f}%)")
        print(f"   个股市值: {stats[3]:,} ({stats[3]/total_samples*100:.2f}%)")
        
        if coverage_pct >= 85:
            print("\n✅ 修复成功！覆盖率达标")
        else:
            print(f"\n⚠️  覆盖率{coverage_pct:.2f}%，需要进一步优化")
            
        cur.close()
        return coverage_pct
        
    except Exception as e:
        print(f"❌ 修复失败: {e}")
        conn.rollback()
        return 0
    finally:
        conn.close()

if __name__ == '__main__':
    quick_fix_pb_pe_data()
```

#### 使用方法
```bash
# 保存脚本为quick_fix_pb_pe.py，然后运行：
cd /Users/davy/github/investment_valuation/price_maintenance_risk_analysis
python ml_training/scripts/quick_fix_pb_pe.py
```

### 方案C: 完善stock_daily_basic数据摄入流程（长期方案）

#### 修复策略
建立标准化的数据同步流程，确保stock_daily_basic数据自动同步到ml_features_wide表

#### 具体实现步骤

**Step 1: 创建数据同步函数**
```python
# 在 ml_training/data/data_sync.py 中创建新文件

import pymysql
import pandas as pd
from datetime import datetime

class StockDailyBasicSync:
    """stock_daily_basic数据同步器"""
    
    def __init__(self):
        self.conn = pymysql.connect(host='127.0.0.1', user='root', password='',
                                   database='investment_valuation', charset='utf8mb4')
    
    def sync_to_ml_features_wide(self, batch_size=1000):
        """同步stock_daily_basic数据到ml_features_wide表"""
        
        print("🔄 开始同步stock_daily_basic数据...")
        
        try:
            # 1. 读取需要更新的样本
            df_needed = pd.read_sql('''
                SELECT id, 股票代码, 报价日
                FROM ml_features_wide
                WHERE 个股PB IS NULL
                ORDER BY 报价日 DESC
            ''', self.conn)
            
            print(f"   需要更新的样本数: {len(df_needed):,}")
            
            if len(df_needed) == 0:
                print("   ✅ 所有样本已有数据，无需更新")
                return 100.0
            
            # 2. 批量更新
            total_updated = 0
            for i in range(0, len(df_needed), batch_size):
                batch = df_needed.iloc[i:i+batch_size]
                
                # 构建股票代码列表
                stock_codes = "','".join(batch['股票代码'].astype(str).tolist())
                quote_dates = "','".join(batch['报价日'].astype(str).tolist())
                
                # 执行更新
                with self.conn.cursor() as cursor:
                    cursor.execute(f'''
                        UPDATE ml_features_wide m
                        INNER JOIN stock_daily_basic sb
                            ON m.股票代码 = sb.stock_code
                            AND DATE_FORMAT(m.报价日, "%%Y%%m%%d") = sb.trade_date
                        SET m.个股PB = sb.pb,
                            m.个股PE = sb.pe,
                            m.个股PS = sb.ps,
                            m.个股市值 = sb.total_mv
                        WHERE m.个股PB IS NULL
                          AND m.股票代码 IN ('{stock_codes}')
                    ''')
                    updated = cursor.rowcount
                    self.conn.commit()
                    total_updated += updated
                    
                    print(f"   批次 {i//batch_size + 1}: {updated:,}条更新")
            
            # 3. 验证结果
            with self.conn.cursor() as cursor:
                cursor.execute("SELECT COUNT(*) FROM ml_features_wide WHERE 个股PB IS NOT NULL")
                final_coverage = cursor.fetchone()[0]
                
            cursor.execute("SELECT COUNT(*) FROM ml_features_wide")
            total_samples = cursor.fetchone()[0]
            
            coverage_pct = final_coverage / total_samples * 100
            print(f"\n✅ 同步完成！覆盖率: {final_coverage:,}/{total_samples:,} ({coverage_pct:.2f}%)")
            
            return coverage_pct
            
        except Exception as e:
            print(f"❌ 同步失败: {e}")
            self.conn.rollback()
            return 0
        finally:
            self.conn.close()

# 使用示例
if __name__ == '__main__':
    syncer = StockDailyBasicSync()
    syncer.sync_to_ml_features_wide()
```

**Step 2: 集成到现有流程**
```python
# 在 export_features.py 中添加同步功能

def sync_stock_daily_basic_data():
    """同步stock_daily_basic数据到ml_features_wide表"""
    from ml_training.data.data_sync import StockDailyBasicSync
    
    syncer = StockDailyBasicSync()
    coverage = syncer.sync_to_ml_features_wide()
    
    if coverage >= 85:
        print("✅ 数据同步成功")
    else:
        print(f"⚠️  数据覆盖率{coverage:.2f}%，需要检查")
    
    return coverage

# 在main函数中添加
if __name__ == '__main__':
    # ... 现有代码 ...
    
    # 新增：数据同步步骤
    if args.sync_basic_data:
        sync_stock_daily_basic_data()
    
    # ... 其余代码 ...
```

## 🎯 推荐执行顺序

### 立即执行（今天）
```bash
# 1. 备份数据
mysqldump -u root investment_valuation ml_features_wide > ml_features_wide_backup_$(date +%Y%m%d).sql

# 2. 快速修复（方案B）
python ml_training/scripts/quick_fix_pb_pe.py

# 3. 验证结果
mysql -u root investment_valuation --default-character-set=utf8mb4 -e "
SELECT 
    COUNT(个股PB) as pb_coverage, 
    ROUND(COUNT(个股PB)/COUNT(*)*100,2) as pb_pct,
    COUNT(个股PE) as pe_coverage,
    ROUND(COUNT(个股PE)/COUNT(*)*100,2) as pe_pct
FROM ml_features_wide;"
```

### 本周执行（修复根因）
```bash
# 修改export_features.py（方案A）
# 1. 找到第455行附近的SQL查询
# 2. 添加LEFT JOIN stock_daily_basic
# 3. 修改字段映射逻辑

# 验证修改效果
python ml_training/features/export_features.py --db-register --tag test_pb_fix
```

### 长期优化（建立标准流程）
```bash
# 实施方案C，建立标准数据同步流程
# 创建 ml_training/data/data_sync.py
# 集成到export_features.py主流程
```

## 📊 预期效果

### 修复前（当前状态）
- 个股PB覆盖率: 0%
- 个股PE覆盖率: 0%
- 个股PS覆盖率: 0%
- 个股市值覆盖率: 0%

### 修复后（预期效果）
- 个股PB覆盖率: 90%+
- 个股PE覆盖率: 85%+
- 个股PS覆盖率: 85%+
- 个股市值覆盖率: 95%+

### 业务影响
- ✅ 解决数据流断裂问题
- ✅ 为ML训练提供完整的估值特征
- ✅ 建立标准化的数据同步流程
- ✅ 为未来类似问题提供解决方案

---

**制定时间**: 2026-07-05  
**预计修复时间**: 立即执行30分钟 + 本周2小时  
**风险等级**: 低（有备份，可回滚）  
**预期收益**: 核心字段覆盖率从0%提升到90%+