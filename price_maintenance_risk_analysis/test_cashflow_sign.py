#!/usr/bin/env python3
"""
测试Tushare现金流量表数据的符号
验证c_pay_acq_const_fiolta字段的符号是正数还是负数
"""

import tushare as ts
import os
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

def test_cashflow_sign():
    """测试现金流量表字段的符号"""
    print("=" * 60)
    print("测试Tushare现金流量表数据符号")
    print("=" * 60)

    pro = ts.pro_api(os.environ.get('TUSHARE_TOKEN'))

    # 测试多个股票
    test_stocks = ['600000.SH', '000001.SZ', '300735.SZ']

    for ts_code in test_stocks:
        print(f"\n{'='*60}")
        print(f"股票: {ts_code}")
        print(f"{'='*60}")

        try:
            # 获取最新的现金流量表数据
            df = pro.cashflow(
                ts_code=ts_code,
                start_date='20240101',
                end_date='20240630',
                fields='ts_code,end_date,n_cashflow_act,c_pay_acq_const_fiolta'
            )

            if df.empty:
                print(f"  ⚠️ 未获取到数据")
                continue

            # 取最近一期
            row = df.iloc[-1]
            ocf = row['n_cashflow_act']
            capex_raw = row['c_pay_acq_const_fiolta']

            print(f"  报告期: {row['end_date']}")
            print(f"  n_cashflow_act (经营活动现金流净额): {ocf:>15,.0f}")
            print(f"  c_pay_acq_const_fiolta (资本支出原始值): {capex_raw:>15,.0f}")

            # 分析符号
            print(f"\n  符号分析:")
            print(f"    经营活动现金流: {'正数(流入)' if ocf > 0 else '负数(流出)'}")
            print(f"    资本支出原始值: {'正数' if capex_raw > 0 else '负数(现金流出)'}")

            # 当前代码逻辑
            capex_processed = -capex_raw
            fcf_current = ocf - capex_processed

            print(f"\n  当前代码逻辑 (capex = -原始值):")
            print(f"    资本支出处理后: {capex_processed:,.0f}")
            print(f"    FCF = {ocf:,.0f} - {capex_processed:,.0f} = {fcf_current:,.0f}")

            # 如果直接扣除
            fcf_direct = ocf - capex_raw
            print(f"\n  直接扣除原始值 (FCF = OCF - 原始值):")
            print(f"    FCF = {ocf:,.0f} - {capex_raw:,.0f} = {fcf_direct:,.0f}")

            # 判断哪种逻辑正确
            print(f"\n  💡 判断:")
            if capex_raw < 0:
                print(f"    ✅ 原始值为负数（现金流出）")
                print(f"    当前逻辑正确：capex = -({capex_raw:,.0f}) = {capex_processed:,.0f}")
                print(f"    FCF计算: OCF - capex = {ocf:,.0f} - {capex_processed:,.0f} = {fcf_current:,.0f}")
            else:
                print(f"    ⚠️ 原始值为正数")
                print(f"    当前逻辑可能错误：capex = -({capex_raw:,.0f}) = {capex_processed:,.0f}")
                print(f"    建议改为: capex = {capex_raw:,.0f}（直接使用原始值）")
                print(f"    FCF计算: OCF - capex = {ocf:,.0f} - {capex_raw:,.0f} = {fcf_direct:,.0f}")

        except Exception as e:
            print(f"  ❌ 错误: {e}")

    print(f"\n{'='*60}")
    print("测试完成")
    print("=" * 60)

if __name__ == '__main__':
    test_cashflow_sign()
