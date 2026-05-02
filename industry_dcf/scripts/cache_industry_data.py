#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pre-fetch and cache financial data for an L3 industry.

Usage:
    python -m industry_dcf.scripts.cache_industry_data --l3 850111.SI
    python -m industry_dcf.scripts.cache_industry_data --stock 002001.SZ
"""

import argparse
import os
import sys

import tushare as ts

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from industry_dcf.utils.rate_limiter import RateLimiter
from industry_dcf.utils.shenwan_lookup import find_l3_by_code
from industry_dcf.utils.industry_data_fetcher import IndustryDataFetcher


def main():
    parser = argparse.ArgumentParser(description='缓存行业财务数据')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--l3', help='申万三级行业代码')
    group.add_argument('--stock', help='股票代码 (自动查找行业)')
    parser.add_argument('--years', type=int, default=10, help='获取年数')
    parser.add_argument('--token', default=None, help='Tushare Token')
    args = parser.parse_args()

    token = args.token or os.environ.get('TUSHARE_TOKEN', '')
    if not token:
        print("错误: 请设置 TUSHARE_TOKEN 环境变量")
        sys.exit(1)

    ts.set_token(token)
    pro = ts.pro_api()

    l3_code = args.l3
    if not l3_code:
        info = find_l3_by_code(args.stock, pro)
        if not info:
            print(f"错误: 无法找到 {args.stock} 的行业分类")
            sys.exit(1)
        l3_code = info['l3_code']
        print(f"行业: {info['l1_name']} > {info['l2_name']} > {info['l3_name']}")

    fetcher = IndustryDataFetcher(pro)
    result = fetcher.get_industry_financials(l3_code, years=args.years)
    print(f"完成: 缓存了 {result['company_count']} 家公司的数据")


if __name__ == '__main__':
    main()
