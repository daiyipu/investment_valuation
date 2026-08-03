#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
改进的PB/PE回填脚本 - 支持最近交易日匹配

策略改进：
1. 优先精确匹配报价日与交易日
2. 若无精确匹配，向前查找最近交易日（最多向前5天）
3. 使用更小的批次大小避免锁超时
4. 单条更新减少锁竞争
"""

import pymysql
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys
import os
import time

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, PKG)

def get_nearest_trading_data(cursor, stock_code, quote_date_str, max_lookback=5):
    """向前查找最近交易日的PB/PE数据"""

    try:
        # 处理不同类型的日期输入
        if isinstance(quote_date_str, str):
            formatted_date = quote_date_str
            quote_date = datetime.strptime(quote_date_str, '%Y-%m-%d')
        elif hasattr(quote_date_str, 'strftime'):
            # 如果是date或datetime对象
            formatted_date = quote_date_str.strftime('%Y-%m-%d')
            quote_date = datetime.strptime(formatted_date, '%Y-%m-%d')
        else:
            return None
    except Exception as e:
        return None

    # 尝试精确匹配
    cursor.execute("""
        SELECT pb, pe, ps, total_mv, trade_date
        FROM stock_daily_basic
        WHERE stock_code = %s
          AND trade_date = DATE_FORMAT(%s, '%%Y%%m%%d')
        LIMIT 1
    """, (stock_code, formatted_date))

    result = cursor.fetchone()
    if result and result[0] is not None:
        return result  # 精确匹配成功

    # 向前查找最近交易日
    for days_back in range(1, max_lookback + 1):
        prior_date = quote_date - timedelta(days=days_back)
        prior_date_str = prior_date.strftime('%Y-%m-%d')

        cursor.execute("""
            SELECT pb, pe, ps, total_mv, trade_date
            FROM stock_daily_basic
            WHERE stock_code = %s
              AND trade_date = DATE_FORMAT(%s, '%%Y%%m%%d')
            LIMIT 1
        """, (stock_code, prior_date_str))

        result = cursor.fetchone()
        if result and result[0] is not None:
            return result  # 找到最近交易日数据

    return None  # 未找到任何数据

def backfill_with_nearest_trading_day(batch_size=1000):
    """使用最近交易日策略回填PB/PE"""
    print(f"🔄 最近交易日策略回填PB/PE（批次: {batch_size}）...")

    conn = pymysql.connect(
        host='127.0.0.1',
        user='root',
        password='',
        database='investment_valuation',
        charset='utf8mb4'
    )

    try:
        cursor = conn.cursor()

        # 检查现有情况
        print("📊 检查现有PB、PE特征缺失情况...")
        cursor.execute("""
            SELECT
                COUNT(*) as total,
                SUM(CASE WHEN 个股PB IS NULL THEN 1 ELSE 0 END) as missing_pb,
                SUM(CASE WHEN 个股PE IS NULL THEN 1 ELSE 0 END) as missing_pe
            FROM ml_features_wide
        """)

        total, missing_pb, missing_pe = cursor.fetchone()
        print(f"  总样本: {total:,}")
        print(f"  缺失PB: {missing_pb:,} ({missing_pb/total*100:.1f}%)")
        print(f"  缺失PE: {missing_pe:,} ({missing_pe/total*100:.1f}%)")

        # 获取需要更新的样本
        print(f"🔄 开始分批更新...")
        cursor.execute("""
            SELECT id, 股票代码, 报价日
            FROM ml_features_wide
            WHERE 个股PB IS NULL OR 个股PE IS NULL
            ORDER BY id
            LIMIT %s
        """, (batch_size * 100,))

        samples_to_update = cursor.fetchall()
        total_samples = len(samples_to_update)
        print(f"  本次处理样本: {total_samples:,}")

        if total_samples == 0:
            print("  ✅ 所有样本已有数据，无需更新")
            return

        # 统计数据
        exact_matches = 0
        nearest_matches = 0
        no_matches = 0
        total_updated = 0
        batch_num = 0

        # 分批处理
        for i in range(0, total_samples, batch_size):
            batch_num += 1
            batch_samples = samples_to_update[i:i+batch_size]
            end_i = min(i+batch_size, total_samples)

            print(f"  批次 {batch_num}: 处理 {len(batch_samples):,} 条样本 ({i+1}-{end_i})...")

            # 记录批次开始时的统计
            exact_matches_start = exact_matches
            nearest_matches_start = nearest_matches
            no_matches_start = no_matches

            for sample_id, stock_code, quote_date in batch_samples:
                try:
                    # 获取最近交易日数据
                    trading_data = get_nearest_trading_data(cursor, stock_code, quote_date)

                    # 调试信息（仅前3个样本）
                    if exact_matches + nearest_matches + no_matches < 3:
                        print(f"    DEBUG: {stock_code} {quote_date} (type: {type(quote_date).__name__}) -> {trading_data}")

                    if trading_data:
                        pb, pe, ps, total_mv, trade_date = trading_data

                        # 更新数据库
                        cursor.execute("""
                            UPDATE ml_features_wide
                            SET 个股PB = %s,
                                个股PE = %s,
                                个股PS = %s,
                                个股市值 = %s
                            WHERE id = %s
                        """, (pb, pe, ps, total_mv, sample_id))

                        total_updated += cursor.rowcount

                        # 统计匹配类型
                        quote_date_formatted = datetime.strptime(quote_date, '%Y-%m-%d').strftime('%Y%m%d')
                        if trade_date == quote_date_formatted:
                            exact_matches += 1
                        else:
                            nearest_matches += 1
                    else:
                        no_matches += 1

                except Exception as e:
                    print(f"    ⚠️ 样本 {sample_id} 更新失败: {e}")
                    continue

            # 提交当前批次
            conn.commit()

            batch_exact = exact_matches - exact_matches_start
            batch_nearest = nearest_matches - nearest_matches_start
            batch_no_match = no_matches - no_matches_start

            print(f"    本批次: 精确匹配 {batch_exact}, 最近交易日 {batch_nearest}, 无匹配 {batch_no_match}")
            print(f"    累计统计: 精确匹配 {exact_matches}, 最近交易日 {nearest_matches}, 无匹配 {no_matches}")

            # 短暂暂停，避免持续锁竞争
            if batch_num % 5 == 0:
                time.sleep(0.5)

        print(f"✅ PB/PE批量更新完成！总计更新 {total_updated:,}条样本")

        # 最终统计
        print(f"\\n📊 匹配统计:")
        print(f"   精确匹配: {exact_matches:,} ({exact_matches/total_samples*100:.1f}%)")
        print(f"   最近交易日匹配: {nearest_matches:,} ({nearest_matches/total_samples*100:.1f}%)")
        print(f"   无匹配: {no_matches:,} ({no_matches/total_samples*100:.1f}%)")
        print(f"   总覆盖率: {(exact_matches + nearest_matches)/total_samples*100:.1f}%")

        # 最终检查覆盖率
        cursor.execute("""
            SELECT
                COUNT(*) as total,
                SUM(CASE WHEN 个股PB IS NOT NULL THEN 1 ELSE 0 END) as has_pb,
                SUM(CASE WHEN 个股PE IS NOT NULL THEN 1 ELSE 0 END) as has_pe,
                SUM(CASE WHEN 个股PS IS NOT NULL THEN 1 ELSE 0 END) as has_ps,
                SUM(CASE WHEN 个股市值 IS NOT NULL THEN 1 ELSE 0 END) as has_mv
            FROM ml_features_wide
        """)

        result = cursor.fetchone()
        total = result[0]
        has_pb = result[1]
        has_pe = result[2]
        has_ps = result[3]
        has_mv = result[4]

        print(f"\\n📊 最终覆盖率:")
        print(f"   个股PB: {has_pb:,}/{total:,} ({has_pb/total*100:.1f}%)")
        print(f"   个股PE: {has_pe:,}/{total:,} ({has_pe/total*100:.1f}%)")
        print(f"   个股PS: {has_ps:,}/{total:,} ({has_ps/total*100:.1f}%)")
        print(f"   个股市值: {has_mv:,}/{total:,} ({has_mv/total*100:.1f}%)")

    except Exception as e:
        print(f"❌ 更新失败: {e}")
        import traceback
        traceback.print_exc()
        conn.rollback()
    finally:
        cursor.close()
        conn.close()

def main():
    """主函数"""
    print("🚀 最近交易日策略回填PB/PE")
    print("=" * 50)

    backfill_with_nearest_trading_day(batch_size=1000)

    print("\n🎉 PB/PE批量更新完成!")

if __name__ == '__main__':
    main()