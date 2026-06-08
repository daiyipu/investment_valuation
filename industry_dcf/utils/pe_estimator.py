#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Normalized PE Estimation Module.

Uses industry profitability benchmarks to estimate "normalized PE" —
what PE would be if the company earned industry-normal profits —
and projects forward PE for 1-10 years based on industry earnings growth.
"""

import os
import sys
from datetime import datetime

import numpy as np
import pandas as pd

_project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from industry_dcf.utils.shenwan_lookup import find_l3_by_code
from industry_dcf.utils.industry_data_fetcher import IndustryDataFetcher
from industry_dcf.utils.rate_limiter import RateLimiter


class PEEstimator:
    """Estimate normalized PE using industry profitability benchmarks."""

    # ---------- public API ----------

    def estimate_normalized_pe(
        self,
        ts_code: str,
        industry_financials: dict,
        industry_benchmark: dict,
        company_financials: dict,
        market_data: dict,
    ) -> dict:
        """Compute normalized PE and forward PE projections.

        Args:
            ts_code: Target stock code.
            industry_financials: Raw industry data from IndustryDataFetcher.
            industry_benchmark: Pre-computed benchmark from calculate_industry_benchmark().
            company_financials: Target company data from get_company_financials().
            market_data: {'market_cap': yuan, 'current_price': yuan, 'total_shares': shares}.

        Returns:
            Dict with normalized PE analysis and forward projections.
        """
        result = {'ts_code': ts_code}

        # --- industry net margin ---
        ind_margin = self._compute_industry_net_margin(industry_financials)
        result['industry_net_margin'] = ind_margin

        # --- industry earnings growth ---
        ind_eg = self._compute_industry_earnings_growth(industry_financials)
        result['industry_earnings_growth'] = ind_eg

        # --- company net margin ---
        co_margin = self._compute_company_net_margin(company_financials)
        result['company_net_margin'] = co_margin

        result['margin_gap'] = co_margin.get('latest', 0) - ind_margin.get('median', 0)

        # --- extract company revenue & net income (wan yuan) ---
        company_rev, company_ni = self._extract_company_revenue_ni(company_financials)
        result['company_revenue_latest'] = company_rev
        result['actual_net_income'] = company_ni

        # --- market data ---
        market_cap_wan = market_data.get('market_cap', 0) / 10000  # yuan -> wan yuan
        total_shares_wan = market_data.get('total_shares', 0)  # already in wan shares from Tushare total_share
        current_price = market_data.get('current_price', 0)
        result['market_cap_wan'] = market_cap_wan
        result['current_price'] = current_price

        # --- actual PE ---
        if company_ni > 0:
            result['actual_pe'] = market_cap_wan / company_ni
        else:
            result['actual_pe'] = None

        # --- normalized PE ---
        # Use the higher of company margin and industry margin
        # If company margin > industry, company is already above-normal, keep actual
        ind_margin_val = ind_margin.get('median', 0)
        co_margin_latest = co_margin.get('latest', 0)
        normalized_margin = max(ind_margin_val, co_margin_latest) if co_margin_latest > 0 else ind_margin_val

        if normalized_margin > 0 and company_rev > 0:
            normalized_ni = company_rev * normalized_margin
            result['normalized_net_income'] = normalized_ni
            result['normalized_margin_used'] = normalized_margin
            result['normalized_margin_source'] = ('company' if co_margin_latest >= ind_margin_val
                                                   else 'industry')
            if normalized_ni > 0:
                result['normalized_pe'] = market_cap_wan / normalized_ni
            else:
                result['normalized_pe'] = None
        else:
            result['normalized_net_income'] = 0
            result['normalized_pe'] = None

        # --- PE gap ---
        actual_pe = result.get('actual_pe')
        norm_pe = result.get('normalized_pe')
        if actual_pe and norm_pe and norm_pe > 0:
            result['pe_gap'] = actual_pe - norm_pe
            result['pe_gap_pct'] = (actual_pe / norm_pe - 1) * 100
        else:
            result['pe_gap'] = None
            result['pe_gap_pct'] = None

        # --- industry PE from benchmark ---
        ind_pe = industry_benchmark.get('industry_pe', {})
        result['industry_pe_median'] = ind_pe.get('median', 0)
        result['industry_pe_stats'] = ind_pe

        # --- forward PE projections ---
        earnings_growth = ind_eg.get('median', 0.10)
        result['earnings_growth_used'] = earnings_growth
        if norm_pe and normalized_ni > 0 and total_shares_wan > 0:
            result['forward_pe_projections'] = self._project_forward_pe(
                market_cap_wan=market_cap_wan,
                normalized_ni=normalized_ni,
                earnings_growth=earnings_growth,
                industry_pe=ind_pe.get('median', 0),
                total_shares_wan=total_shares_wan,
                current_price=current_price,
            )
        else:
            result['forward_pe_projections'] = []

        return result

    # ---------- private helpers ----------

    def _compute_industry_net_margin(self, industry_financials: dict) -> dict:
        """Compute industry net profit margin statistics.

        Uses n_income_attr_p / total_revenue for each company-year.
        Returns median, mean, p25, p75, valid_count.
        """
        margins = []
        companies = industry_financials.get('companies', {})
        for ts_code, data in companies.items():
            inc_records = data.get('income') or []
            if not inc_records:
                continue
            annual = [r for r in inc_records
                      if str(r.get('end_date', '')).endswith('1231')]
            # deduplicate by end_date, keep last
            seen = {}
            for r in annual:
                seen[r['end_date']] = r
            for r in seen.values():
                rev = r.get('total_revenue', 0)
                ni = r.get('n_income_attr_p')
                if ni is None:
                    ni = r.get('n_income', 0)
                if not rev or rev <= 0:
                    continue
                if ni is None:
                    continue
                margin = float(ni) / float(rev)
                if abs(margin) > 1:
                    continue
                margins.append(margin)

        if not margins:
            return {'median': 0, 'mean': 0, 'p25': 0, 'p75': 0, 'valid_count': 0}
        arr = np.array(margins)
        return {
            'median': float(np.median(arr)),
            'mean': float(np.mean(arr)),
            'p25': float(np.percentile(arr, 25)),
            'p75': float(np.percentile(arr, 75)),
            'valid_count': len(margins),
        }

    def _compute_industry_earnings_growth(self, industry_financials: dict) -> dict:
        """Compute industry earnings growth rate statistics.

        Uses YoY growth of n_income_attr_p.
        """
        growth_rates = []
        companies = industry_financials.get('companies', {})
        for ts_code, data in companies.items():
            inc_records = data.get('income') or []
            if not inc_records:
                continue
            annual = [r for r in inc_records
                      if str(r.get('end_date', '')).endswith('1231')]
            seen = {}
            for r in annual:
                seen[r['end_date']] = r
            # build {year: ni} map
            ni_map = {}
            for r in seen.values():
                year = str(r['end_date'])[:4]
                ni = r.get('n_income_attr_p')
                if ni is None:
                    ni = r.get('n_income')
                if ni is not None:
                    ni_map[year] = float(ni)
            # YoY growth
            years = sorted(ni_map.keys())
            for i in range(1, len(years)):
                prev = ni_map[years[i - 1]]
                curr = ni_map[years[i]]
                if prev == 0:
                    continue
                g = (curr - prev) / abs(prev)
                if -0.5 < g < 2.0:
                    growth_rates.append(g)

        if not growth_rates:
            return {'median': 0, 'mean': 0, 'valid_count': 0}
        arr = np.array(growth_rates)
        return {
            'median': float(np.median(arr)),
            'mean': float(np.mean(arr)),
            'valid_count': len(growth_rates),
        }

    def _compute_company_net_margin(self, company_financials: dict) -> dict:
        """Compute target company net profit margin from income statement.

        Uses n_income_attr_p / total_revenue.
        """
        inc_df = company_financials.get('income')
        if inc_df is None or (isinstance(inc_df, pd.DataFrame) and inc_df.empty):
            return {'latest': 0, 'avg_3y': 0, 'by_year': {}}

        # handle both DataFrame and list-of-dict
        if isinstance(inc_df, pd.DataFrame):
            records = inc_df.to_dict('records')
        else:
            records = inc_df

        annual = [r for r in records
                  if str(r.get('end_date', '')).endswith('1231')]
        # deduplicate
        seen = {}
        for r in annual:
            seen[r['end_date']] = r

        by_year = {}
        for r in seen.values():
            year = str(r['end_date'])[:4]
            rev = r.get('total_revenue', 0)
            ni = r.get('n_income_attr_p')
            if ni is None:
                ni = r.get('n_income', 0)
            if rev and float(rev) > 0 and ni is not None:
                by_year[year] = float(ni) / float(rev)

        sorted_years = sorted(by_year.keys())
        latest = by_year.get(sorted_years[-1], 0) if sorted_years else 0
        recent_3 = [by_year[y] for y in sorted_years[-3:]] if len(sorted_years) >= 1 else []
        avg_3y = sum(recent_3) / len(recent_3) if recent_3 else 0

        return {'latest': latest, 'avg_3y': avg_3y, 'by_year': by_year}

    def _extract_company_revenue_ni(self, company_financials: dict):
        """Extract latest revenue and net income (wan yuan) from company data.

        优先用最新年报(end_date含1231)，无年报时回退到最新季报。
        """
        inc_df = company_financials.get('income')
        if inc_df is None or (isinstance(inc_df, pd.DataFrame) and inc_df.empty):
            return 0, 0

        if isinstance(inc_df, pd.DataFrame):
            records = inc_df.to_dict('records')
        else:
            records = inc_df

        if not records:
            return 0, 0

        # 优先年报，无则用最新报告期
        annual = [r for r in records if str(r.get('end_date', '')).endswith('1231')]
        pool = annual if annual else records

        # deduplicate by end_date, keep last
        seen = {}
        for r in pool:
            ed = r.get('end_date', '')
            if ed:
                seen[ed] = r

        if not seen:
            return 0, 0

        latest = seen[sorted(seen.keys())[-1]]
        rev = float(latest.get('total_revenue', 0) or 0) / 10000  # yuan -> wan yuan
        ni = latest.get('n_income_attr_p')
        if ni is None:
            ni = latest.get('n_income', 0)
        ni = float(ni or 0) / 10000

        # 季报年化（若用的是季报且非年报）
        if not annual and str(latest.get('end_date', '')).endswith(('0331', '0630', '0930')):
            end_date = str(latest.get('end_date', ''))
            month = end_date[4:6]
            quarters = {'03': 1, '06': 2, '09': 3}.get(month, 1)
            if quarters > 0 and rev > 0:
                rev = rev / quarters * 4  # 年化营收
                ni = ni / quarters * 4     # 年化净利润

        return rev, ni

    def _project_forward_pe(
        self,
        market_cap_wan: float,
        normalized_ni: float,
        earnings_growth: float,
        industry_pe: float,
        total_shares_wan: float,
        current_price: float,
        n_years: int = 10,
    ) -> list:
        """Project forward PE and target price for each year 1..n."""
        projections = []
        for year in range(1, n_years + 1):
            proj_ni = normalized_ni * ((1 + earnings_growth) ** year)
            fwd_pe = market_cap_wan / proj_ni if proj_ni > 0 else 0
            # target price if market values at industry PE
            proj_price = (proj_ni * industry_pe / total_shares_wan) if (
                industry_pe > 0 and total_shares_wan > 0) else 0
            upside = ((proj_price / current_price) - 1) * 100 if current_price > 0 else 0
            projections.append({
                'year': year,
                'projected_ni': proj_ni,
                'forward_pe': fwd_pe,
                'projected_price': proj_price,
                'upside_pct': upside,
            })
        return projections


# ---------- module-level convenience ----------

def run_pe_estimation(
    ts_code: str,
    tushare_token: str = None,
    force_refresh: bool = False,
    l3_code: str = None,
) -> dict:
    """Standalone PE estimation pipeline."""
    import tushare as ts

    token = tushare_token or os.environ.get('TUSHARE_TOKEN', '')
    if not token:
        return {'error': '请设置 TUSHARE_TOKEN 环境变量'}

    ts.set_token(token)
    pro = ts.pro_api()

    rate_limiter = RateLimiter()
    data_fetcher = IndustryDataFetcher(pro, rate_limiter=rate_limiter)

    # industry lookup
    if l3_code:
        industry_info = {'l3_code': l3_code, 'l3_name': l3_code}
    else:
        industry_info = find_l3_by_code(ts_code, pro)
        if not industry_info:
            return {'error': f'无法找到 {ts_code} 的行业分类'}

    # fetch data
    print("获取行业财务数据...")
    industry_financials = data_fetcher.get_industry_financials(
        industry_info['l3_code'], force_refresh=force_refresh,
    )

    print("获取行业PE数据...")
    industry_pe_data = data_fetcher.get_industry_daily_basics(industry_info['l3_code'])

    # compute benchmark (needed for industry_pe stats)
    from industry_dcf.utils.industry_dcf_calculator import IndustryDCFCalculator
    calculator = IndustryDCFCalculator()
    benchmark = calculator.calculate_industry_benchmark(
        industry_financials, industry_pe_data=industry_pe_data,
    )

    # company data
    print(f"获取 {ts_code} 财务数据...")
    company_data = data_fetcher.get_company_financials(ts_code)
    if company_data is None:
        return {'error': f'无法获取 {ts_code} 的财务数据'}

    # market data
    market_data = _fetch_market_data(ts_code, pro, rate_limiter)

    # estimate
    estimator = PEEstimator()
    result = estimator.estimate_normalized_pe(
        ts_code=ts_code,
        industry_financials=industry_financials,
        industry_benchmark=benchmark,
        company_financials=company_data,
        market_data=market_data,
    )
    result['industry_info'] = industry_info
    return result


def _fetch_market_data(ts_code: str, pro, rate_limiter: RateLimiter) -> dict:
    """Fetch current price, total shares, and market cap."""
    from datetime import timedelta

    market_data = {'current_price': 0, 'total_shares': 0, 'market_cap': 0}
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


def print_pe_result(result: dict):
    """Print formatted PE estimation result."""
    if not result or 'error' in result:
        print(f"  PE估值错误: {result.get('error', '未知')}")
        return

    ts_code = result.get('ts_code', '?')
    print("\n" + "=" * 70)
    print(f"  标准化PE估值报告: {ts_code}")
    print(f"  日期: {datetime.now().strftime('%Y-%m-%d')}")
    print("=" * 70)

    # company profitability
    co_m = result.get('company_net_margin', {})
    ind_m = result.get('industry_net_margin', {})
    margin_gap = result.get('margin_gap', 0)

    print("\n📊 公司盈利能力:")
    print(f"  公司净利率(最新):     {co_m.get('latest', 0)*100:.2f}%")
    print(f"  公司净利率(3年均):    {co_m.get('avg_3y', 0)*100:.2f}%")
    print(f"  行业净利率中位数:     {ind_m.get('median', 0)*100:.2f}%"
          f"  (p25={ind_m.get('p25',0)*100:.1f}%, p75={ind_m.get('p75',0)*100:.1f}%)")
    print(f"  净利率差距:           {margin_gap*100:+.2f}%"
          f"  ({'低于' if margin_gap < 0 else '高于'}行业"
          f" {abs(margin_gap)/ind_m.get('median',0.01)*100:.1f}%)" if ind_m.get('median', 0) != 0 else "")

    # industry growth
    ind_eg = result.get('industry_earnings_growth', {})
    print("\n📈 行业增长参数:")
    print(f"  行业净利润增长率中位数:  {ind_eg.get('median', 0)*100:.2f}%"
          f"  (均值={ind_eg.get('mean',0)*100:.1f}%, {ind_eg.get('valid_count',0)}个样本)")

    # PE analysis
    actual_pe = result.get('actual_pe')
    norm_pe = result.get('normalized_pe')
    ind_pe = result.get('industry_pe_median', 0)
    pe_gap = result.get('pe_gap')

    print("\n💰 PE分析:")
    pe_str = f"{actual_pe:.1f}" if actual_pe else "N/A(亏损)"
    norm_str = f"{norm_pe:.1f}" if norm_pe else "N/A"
    gap_str = f"{pe_gap:+.1f}" if pe_gap is not None else "N/A"
    gap_pct = result.get('pe_gap_pct')
    gap_pct_str = f"({gap_pct:+.1f}%)" if gap_pct is not None else ""

    print(f"  实际PE(静态):         {pe_str}")
    print(f"  标准化PE:             {norm_str}  (假设行业正常盈利)")
    print(f"  行业PE中位数:         {ind_pe:.1f}")
    print(f"  PE差距:               {gap_str} {gap_pct_str}")

    # normalized NI
    rev = result.get('company_revenue_latest', 0)
    ni_actual = result.get('actual_net_income', 0)
    ni_norm = result.get('normalized_net_income', 0)
    norm_src = result.get('normalized_margin_source', 'industry')
    norm_margin = result.get('normalized_margin_used', 0)

    print("\n📋 标准化净利润推算:")
    print(f"  最新营收:             {rev:,.0f} 万元 ({rev/10000:.2f} 亿元)")
    print(f"  实际归母净利润:       {ni_actual:,.0f} 万元 ({ni_actual/10000:.2f} 亿元)")
    src_label = '公司实际' if norm_src == 'company' else '行业'
    print(f"  标准化净利润:         {ni_norm:,.0f} 万元 ({ni_norm/10000:.2f} 亿元)"
          f"  (营收 × {src_label}净利率{norm_margin*100:.1f}%)")

    # forward PE table
    projections = result.get('forward_pe_projections', [])
    if projections:
        eg = result.get('earnings_growth_used', 0)
        print(f"\n🔮 前瞻PE预测 (增长率={eg*100:.1f}%):")
        print(f"  {'年份':>4}  {'预测净利润(万)':>14}  {'前瞻PE':>8}  "
              f"{'行业PE对应价':>12}  {'涨跌幅':>8}")
        print("  " + "-" * 60)
        for p in projections:
            yr = p['year']
            ni = p['projected_ni']
            fpe = p['forward_pe']
            price = p['projected_price']
            up = p['upside_pct']
            # show key years: 1,2,3,5,7,10
            if yr in (1, 2, 3, 5, 7, 10):
                print(f"  {yr:>4}  {ni:>14,.0f}  {fpe:>8.1f}  "
                      f"{price:>10.2f}元  {up:>+7.1f}%")

        print(f"\n  注: 前瞻PE = 当前市值 / (标准化净利润 × (1+{eg*100:.1f}%)^年数)")
        if ind_pe > 0:
            print(f"      行业PE对应股价 = 预测净利润 × 行业PE中位数({ind_pe:.1f}) / 总股本")

    print("\n" + "=" * 70)
