#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Industry-Calibrated DCF Valuation - Entry Point.

Usage:
    python -m industry_dcf.main --stock 002001.SZ
    python -m industry_dcf.main --stock 002001.SZ --force-refresh
    python -m industry_dcf.main --l3 850111.SI --stock 002001.SZ
"""

import argparse
import os
import sys

import tushare as ts

_project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from industry_dcf.utils.rate_limiter import RateLimiter
from industry_dcf.utils.shenwan_lookup import find_l3_by_code
from industry_dcf.utils.industry_data_fetcher import IndustryDataFetcher
from industry_dcf.utils.industry_dcf_calculator import IndustryDCFCalculator

# Try to import WACCCalculator from valuation_report
try:
    from valuation_report.utils.wacc_calculator import WACCCalculator
except ImportError:
    WACCCalculator = None


def run_industry_dcf(
    ts_code: str,
    tushare_token: str = None,
    force_refresh: bool = False,
    l3_code: str = None,
    params: dict = None,
) -> dict:
    """Full industry-calibrated DCF pipeline.

    Args:
        ts_code: Target stock code (e.g., '002001.SZ').
        tushare_token: Tushare API token. If None, reads from TUSHARE_TOKEN env var.
        force_refresh: Force re-fetch industry data.
        l3_code: Override L3 industry code (skip auto-detection).
        params: Override default DCF parameters.

    Returns:
        Full valuation result dict.
    """
    # 1. Initialize Tushare
    token = tushare_token or os.environ.get('TUSHARE_TOKEN', '')
    if not token:
        return {'error': '请设置 TUSHARE_TOKEN 环境变量或传入 tushare_token 参数'}

    ts.set_token(token)
    pro = ts.pro_api()

    rate_limiter = RateLimiter()
    data_fetcher = IndustryDataFetcher(pro, rate_limiter=rate_limiter)
    wacc_calc = WACCCalculator(pro) if WACCCalculator else None
    calculator = IndustryDCFCalculator(wacc_calc=wacc_calc)

    # 2. Find industry
    if l3_code:
        industry_info = {'l3_code': l3_code, 'l3_name': l3_code}
        print(f"使用指定行业: {l3_code}")
    else:
        print(f"查询 {ts_code} 的申万三级行业...")
        industry_info = find_l3_by_code(ts_code, pro)
        if not industry_info:
            return {'error': f'无法找到 {ts_code} 的行业分类'}
        print(f"  行业: {industry_info['l1_name']} > {industry_info['l2_name']} > {industry_info['l3_name']}")

    # 3. Fetch industry data
    print("获取行业财务数据...")
    industry_financials = data_fetcher.get_industry_financials(
        industry_info['l3_code'], force_refresh=force_refresh,
    )

    # 3b. Fetch industry PE data from daily_basic
    print("获取行业PE数据...")
    industry_pe_data = data_fetcher.get_industry_daily_basics(industry_info['l3_code'])

    # 4. Calculate industry benchmark
    print("计算行业基准参数...")
    benchmark = calculator.calculate_industry_benchmark(
        industry_financials, industry_pe_data=industry_pe_data,
    )
    if 'error' in benchmark:
        return benchmark
    print(f"  行业FCFF/营收中位数: {benchmark['fcff_rev_ratio']['median']:.4f}")
    print(f"  行业FCFF增长率中位数: {benchmark['fcff_growth_rate']['median']:.2%}")
    pe_info = benchmark.get('industry_pe', {})
    if pe_info.get('median'):
        print(f"  行业PE中位数: {pe_info['median']:.1f} ({pe_info.get('valid_count', 0)}家公司)")
    print(f"  隐含预测年数(最新): {benchmark['recommended_forecast_years']}年")

    # Regression info
    reg = benchmark.get('regression', {})
    print(f"    方法: {reg.get('method', 'N/A')} (样本量={reg.get('sample_size', 0)}, R²={reg.get('r_squared', 0)})")
    if reg.get('coefficients'):
        coef = reg['coefficients']
        terms = [f"{coef.get('intercept', 0):.2f}"]
        if 'pe' in coef:
            terms.append(f"{coef['pe']:.4f}×PE")
        terms.append(f"{coef.get('fcff_g', 0):.2f}×FCFF增长率")
        terms.append(f"{coef.get('fcff_rev', 0):.2f}×FCFF/营收")
        terms.append(f"{coef.get('rev_g', 0):.2f}×营收增长率")
        print(f"    回归: Y = {' + '.join(terms)}")

    # Per-year forecast years
    by_year = reg.get('by_year', {})
    if by_year:
        print(f"    逐年预测年数:")
        for year in sorted(by_year.keys()):
            print(f"      {year}: {by_year[year]}年")

    print(f"  行业成熟度: {benchmark['industry_maturity']}")

    # 5. Fetch target company data
    print(f"获取 {ts_code} 财务数据...")
    company_data = data_fetcher.get_company_financials(ts_code)
    if company_data is None:
        return {'error': f'无法获取 {ts_code} 的财务数据'}

    # 6. Market data (current price, shares, market cap)
    market_data = _fetch_market_data(ts_code, pro, rate_limiter)

    # 7. Valuation
    print("执行行业校准DCF估值...")
    result = calculator.calculate_company_valuation(
        ts_code=ts_code,
        industry_benchmark=benchmark,
        company_financials=company_data,
        market_data=market_data,
        params=params,
    )

    return result


def _fetch_market_data(ts_code: str, pro, rate_limiter: RateLimiter) -> dict:
    """Fetch current price, total shares, and market cap."""
    from datetime import datetime, timedelta

    market_data = {'current_price': 0, 'total_shares': 0, 'market_cap': 0, 'total_debt': 0}

    # Try daily_basic for market cap and PE
    for days_back in range(0, 10):
        td = (datetime.now() - timedelta(days=days_back)).strftime('%Y%m%d')
        db = rate_limiter.call(
            pro.daily_basic,
            ts_code=ts_code, trade_date=td,
            fields='ts_code,trade_date,total_mv,total_share,close',
        )
        if db is not None and not db.empty:
            row = db.iloc[0]
            market_data['market_cap'] = float(row.get('total_mv', 0) or 0) * 10000
            market_data['total_shares'] = float(row.get('total_share', 0) or 0)
            market_data['current_price'] = float(row.get('close', 0) or 0)
            break

    return market_data


def print_result(result: dict):
    """Pretty-print valuation result."""
    if 'error' in result:
        print(f"\n❌ 错误: {result['error']}")
        return

    print("\n" + "=" * 70)
    print(f"  行业校准DCF估值报告: {result['ts_code']}")
    print(f"  日期: {result['valuation_date']}")
    print("=" * 70)

    print(f"\n📊 公司分析:")
    print(f"  公司FCFF/营收比:     {result.get('company_fcff_rev_ratio', 'N/A')}")
    print(f"  行业FCFF/营收中位数:  {result['industry_fcff_rev_ratio_median']:.4f}")
    print(f"  调整系数 α:          {result['alpha']:.4f}")
    print(f"  最新营收(万元):      {result['company_revenue_latest']:,.0f}")
    print(f"  公司营收增长率:       {result.get('company_revenue_growth', 0):.2%}")

    print(f"\n📈 增长参数:")
    print(f"  行业增长率:          {result['industry_growth_rate']:.2%}")
    print(f"  混合增长率:          {result['blended_growth_rate']:.2%}")
    print(f"  永续增长率:          {result['terminal_growth_rate']:.2%}")
    print(f"  预测年数:            {result['forecast_years']}年 (由行业PE+FCFF反推)")

    # Industry PE info
    ind_pe = result.get('industry_benchmark', {}).get('industry_pe', {})
    if ind_pe.get('median'):
        print(f"  行业PE中位数:        {ind_pe['median']:.1f} ({ind_pe.get('valid_count', 0)}家公司)")
        print(f"  行业PE P25-P75:      {ind_pe.get('p25', 0):.1f} - {ind_pe.get('p75', 0):.1f}")

    print(f"\n💰 WACC:")
    print(f"  WACC:                {result['wacc']:.2%}")
    print(f"  Ke (股权成本):       {result['ke']:.2%}")
    print(f"  Kd(税后):            {result['kd_aftertax']:.2%}")
    print(f"  Beta:                {result['beta']:.2f}")

    print(f"\n📋 DCF结果 (万元):")
    print(f"  预测期现值:          {result['pv_forecasts']:,.0f}")
    print(f"  终值:                {result['terminal_value']:,.0f}")
    print(f"  终值现值:            {result['pv_terminal']:,.0f}")
    print(f"  企业价值(EV):        {result['enterprise_value']:,.0f}")
    print(f"  净债务:              {result['net_debt']:,.0f}")
    print(f"  股权价值:            {result['equity_value']:,.0f}")

    print(f"\n🎯 估值结论:")
    print(f"  每股价值:            {result['per_share_value']:.2f} 元")
    print(f"  当前股价:            {result['current_price']:.2f} 元")
    margin = result['safety_margin_pct']
    direction = "↑" if margin > 0 else "↓"
    print(f"  安全边际:            {direction} {abs(margin):.1f}%")

    if result['per_share_value'] > result['current_price']:
        print(f"  >>> 低估，建议买入 <<<")
    else:
        print(f"  >>> 高估，谨慎对待 <<<")

    # Sensitivity summary
    sens = result.get('sensitivity', {})
    if sens:
        print(f"\n📊 敏感性分析 (每股价值):")
        wacc_vals = sens['wacc_values']
        growth_vals = sens['growth_values']
        matrix = sens['matrix']

        header = "WACC\\增长率"
        for g in growth_vals:
            header += f"  {g*100:>5.0f}%"
        print(f"  {header}")
        print(f"  {'-' * len(header)}")
        for i, w in enumerate(wacc_vals):
            row_str = f"  {w*100:>7.0f}%"
            for j, g in enumerate(growth_vals):
                cell = matrix[i][j]
                row_str += f"  {cell['per_share']:>6.2f}"
            print(row_str)

    print("\n" + "=" * 70)


def main():
    parser = argparse.ArgumentParser(description='行业校准DCF估值')
    parser.add_argument('--stock', required=True, help='股票代码 (如 002001.SZ)')
    parser.add_argument('--token', default=None, help='Tushare API Token')
    parser.add_argument('--l3', default=None, help='指定申万三级行业代码')
    parser.add_argument('--force-refresh', action='store_true', help='强制刷新缓存')
    parser.add_argument('--forecast-years', type=int, default=None, help='预测年数')
    parser.add_argument('--terminal-growth', type=float, default=None, help='永续增长率')
    parser.add_argument('--industry-weight', type=float, default=None, help='行业权重(0-1)')
    args = parser.parse_args()

    params = {}
    if args.forecast_years:
        params['forecast_years'] = args.forecast_years
    if args.terminal_growth:
        params['terminal_growth_rate'] = args.terminal_growth
    if args.industry_weight:
        params['industry_weight'] = args.industry_weight

    result = run_industry_dcf(
        ts_code=args.stock,
        tushare_token=args.token,
        force_refresh=args.force_refresh,
        l3_code=args.l3,
        params=params or None,
    )

    print_result(result)


if __name__ == '__main__':
    main()
