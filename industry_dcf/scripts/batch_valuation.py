#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Batch valuation across multiple stocks or an entire industry.

Usage:
    python -m industry_dcf.scripts.batch_valuation --file stocks.txt
    python -m industry_dcf.scripts.batch_valuation --l3 850111.SI
"""

import argparse
import os
import sys

import pandas as pd
import tushare as ts

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from industry_dcf.utils.rate_limiter import RateLimiter
from industry_dcf.utils.shenwan_lookup import find_l3_by_code, get_l3_members
from industry_dcf.utils.industry_data_fetcher import IndustryDataFetcher
from industry_dcf.utils.industry_dcf_calculator import IndustryDCFCalculator
from industry_dcf.main import _fetch_market_data

try:
    from valuation_report.utils.wacc_calculator import WACCCalculator
except ImportError:
    WACCCalculator = None


def main():
    parser = argparse.ArgumentParser(description='批量行业校准DCF估值')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--file', help='股票代码文件 (一行一个)')
    group.add_argument('--l3', help='申万三级行业代码 (估值整个行业)')
    parser.add_argument('--token', default=None, help='Tushare Token')
    parser.add_argument('--output', default=None, help='输出CSV文件路径')
    args = parser.parse_args()

    token = args.token or os.environ.get('TUSHARE_TOKEN', '')
    if not token:
        print("错误: 请设置 TUSHARE_TOKEN 环境变量")
        sys.exit(1)

    ts.set_token(token)
    pro = ts.pro_api()

    # Get stock list
    if args.file:
        with open(args.file, 'r') as f:
            stocks = [line.strip() for line in f if line.strip()]
    else:
        stocks = get_l3_members(args.l3, pro)

    if not stocks:
        print("错误: 无股票")
        sys.exit(1)

    print(f"共 {len(stocks)} 只股票待估值")

    rate_limiter = RateLimiter()
    data_fetcher = IndustryDataFetcher(pro, rate_limiter=rate_limiter)
    wacc_calc = WACCCalculator(pro) if WACCCalculator else None
    calculator = IndustryDCFCalculator(wacc_calc=wacc_calc)

    results = []
    for i, ts_code in enumerate(stocks):
        print(f"\n[{i+1}/{len(stocks)}] {ts_code}")

        try:
            # Find industry
            industry_info = find_l3_by_code(ts_code, pro)
            if not industry_info:
                print(f"  跳过: 无法找到行业分类")
                continue

            # Fetch industry data (cached after first fetch)
            industry_financials = data_fetcher.get_industry_financials(industry_info['l3_code'])

            # Benchmark
            benchmark = calculator.calculate_industry_benchmark(industry_financials)
            if 'error' in benchmark:
                print(f"  跳过: {benchmark['error']}")
                continue

            # Company data
            company_data = data_fetcher.get_company_financials(ts_code)
            if company_data is None:
                print(f"  跳过: 无财务数据")
                continue

            # Market data
            market_data = _fetch_market_data(ts_code, pro, rate_limiter)

            # Valuation
            result = calculator.calculate_company_valuation(
                ts_code=ts_code,
                industry_benchmark=benchmark,
                company_financials=company_data,
                market_data=market_data,
            )

            if 'error' in result:
                print(f"  跳过: {result['error']}")
                continue

            results.append({
                'ts_code': ts_code,
                'per_share_value': result['per_share_value'],
                'current_price': result['current_price'],
                'safety_margin_pct': result['safety_margin_pct'],
                'alpha': result['alpha'],
                'blended_growth': result['blended_growth_rate'],
                'wacc': result['wacc'],
                'industry': industry_info.get('l3_name', ''),
            })

            margin = result['safety_margin_pct']
            direction = "低估" if margin > 0 else "高估"
            print(f"  每股价值: {result['per_share_value']:.2f} | 股价: {result['current_price']:.2f} | {direction} {abs(margin):.1f}%")

        except Exception as e:
            print(f"  异常: {e}")
            continue

    if not results:
        print("\n无有效估值结果")
        return

    # Summary
    df = pd.DataFrame(results)
    df = df.sort_values('safety_margin_pct', ascending=False)

    print("\n" + "=" * 80)
    print("批量估值汇总:")
    print("=" * 80)
    print(df.to_string(index=False))

    # Save
    output_path = args.output or 'batch_valuation_results.csv'
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"\n结果已保存: {output_path}")


if __name__ == '__main__':
    main()
