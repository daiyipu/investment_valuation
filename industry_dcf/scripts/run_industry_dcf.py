#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CLI script: Run industry-calibrated DCF for a single stock.

Usage:
    python -m industry_dcf.scripts.run_industry_dcf 002001.SZ
    python -m industry_dcf.scripts.run_industry_dcf 002001.SZ --force-refresh
"""

import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..')))

from industry_dcf.main import run_industry_dcf, print_result


def main():
    if len(sys.argv) < 2:
        print("用法: python -m industry_dcf.scripts.run_industry_dcf <股票代码> [--force-refresh]")
        print("示例: python -m industry_dcf.scripts.run_industry_dcf 002001.SZ")
        sys.exit(1)

    ts_code = sys.argv[1]
    force_refresh = '--force-refresh' in sys.argv

    result = run_industry_dcf(ts_code=ts_code, force_refresh=force_refresh)
    print_result(result)


if __name__ == '__main__':
    main()
