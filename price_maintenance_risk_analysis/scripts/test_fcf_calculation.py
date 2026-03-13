#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
测试FCF计算 - 使用修复后的公式
"""

import sys
import os

# 添加父目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from update_market_data import TushareFinancialData

def main():
    print("="*70)
    print("🧪 测试修复后的FCF计算")
    print("="*70)

    stock_code = '300735.SZ'

    try:
        # 创建财务数据获取对象
        financial = TushareFinancialData(stock_code)

        # 获取历史FCF数据
        print(f"\n📊 获取 {stock_code} 的历史FCF数据...")
        historical_fcf = financial.get_historical_fcf_for_dcf(max_years=15)

        if historical_fcf and historical_fcf.get('data'):
            print("\n" + "="*70)
            print("📈 历史FCF计算结果")
            print("="*70)

            data = historical_fcf['data']

            print(f"\n{'年份':<6} {'营收(亿)':<12} {'NOPAT(亿)':<12} {'折旧(亿)':<12} {'CapEx(亿)':<12} {'WC变化(亿)':<12} {'FCF(亿)':<12}")
            print("-"*90)

            for year_data in data[-10:]:  # 显示最近10年
                year = year_data['year']
                revenue = year_data.get('revenue', 0) / 100000000
                nopat = year_data.get('nopat', 0) / 100000000
                depreciation = year_data.get('depreciation', 0) / 100000000
                capex = year_data.get('capex', 0) / 100000000
                wc_change = year_data.get('wc_change', 0) / 100000000
                fcf = year_data['fcf'] / 100000000

                print(f"{year:<6} {revenue:>10.2f}     {nopat:>10.2f}     {depreciation:>10.2f}     {capex:>10.2f}     {wc_change:>10.2f}     {fcf:>10.2f}")

            # 统计信息
            positive_fcf_years = sum(1 for d in data if d['fcf'] > 0)
            total_years = len(data)

            print("\n" + "="*70)
            print("📊 统计信息")
            print("="*70)
            print(f"总年数: {total_years}")
            print(f"FCF为正的年数: {positive_fcf_years} ({positive_fcf_years/total_years*100:.1f}%)")
            print(f"最近3年平均FCF: {sum(d['fcf'] for d in data[-3:])/(3*100000000):.2f} 亿元")

            # 检查最新FCF
            latest_fcf = data[-1]['fcf'] / 100000000
            print(f"\n最新年度FCF: {latest_fcf:.2f} 亿元")

            if latest_fcf > 0:
                print("✅ 最新FCF为正值！")
            else:
                print("⚠️ 最新FCF仍为负值，需要进一步调整")

        else:
            print("❌ 未获取到历史FCF数据")

    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
