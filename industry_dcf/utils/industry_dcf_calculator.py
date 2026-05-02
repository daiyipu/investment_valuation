#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Industry-Calibrated DCF Calculator.

Uses Shenwan L3 industry median parameters as a stable baseline,
then applies individual company adjustment coefficients.

Core formula:
    FCFF_t = Revenue_t × industry_median_fcff_rev_ratio × α
    α = clamp(company_fcff_rev_ratio / industry_median, 0.3, 3.0)

Three-step process:
    1. Calculate industry benchmark (median FCFF/revenue, growth, PE)
    2. Compute individual company adjustment coefficient
    3. Run DCF with blended parameters
"""

import sys
import os
import numpy as np
import pandas as pd
from typing import Dict, Optional

# Import existing DCF/WACC calculators
_project_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from valuation_report.utils.wacc_calculator import WACCCalculator


class IndustryDCFCalculator:
    """Industry-calibrated DCF valuation."""

    ALPHA_MIN = 0.3
    ALPHA_MAX = 3.0

    DEFAULT_PARAMS = {
        'forecast_years': 5,
        'terminal_growth_rate': 0.025,
        'industry_weight': 0.7,
        'alpha_bounds': (0.3, 3.0),
        'risk_free_rate': 0.0185,
        'market_risk_premium': 0.0615,
        'tax_rate': 0.25,
        'kd_pretax': 0.0278,
    }

    def __init__(self, wacc_calc: WACCCalculator = None):
        self.wacc_calc = wacc_calc
        self.params = dict(self.DEFAULT_PARAMS)

    # ================================================================
    # Step 1: Industry Benchmark
    # ================================================================

    def calculate_industry_benchmark(
        self,
        industry_financials: dict,
        industry_pe_data: pd.DataFrame = None,
    ) -> dict:
        """Calculate industry-level benchmark parameters.

        Args:
            industry_financials: Output from IndustryDataFetcher.get_industry_financials()
            industry_pe_data: DataFrame from IndustryDataFetcher.get_industry_daily_basics(),
                              containing 'pe' column for all industry members.

        Returns:
            Industry benchmark dict with median ratios, growth rates, PE, maturity.
        """
        companies = industry_financials.get('companies', {})
        l3_code = industry_financials.get('l3_code', '')

        # Collect per-company FCFF and revenue data
        all_fcff_rev_ratios = []  # all valid FCFF/revenue ratios across companies/years
        all_fcff_growth_rates = []
        all_rev_growth_rates = []
        all_fcff_ni_ratios = []   # FCFF / net_income ratios for PE→EV/FCFF conversion
        company_fcff_series = {}  # ts_code -> {year: fcff}
        company_rev_series = {}   # ts_code -> {year: revenue}
        company_ni_series = {}    # ts_code -> {year: net_income}

        for ts_code, data in companies.items():
            cf_df = pd.DataFrame(data.get('cashflow') or [])
            inc_df = pd.DataFrame(data.get('income') or [])

            if cf_df.empty:
                continue

            fcff_list = self._calc_company_fcff(cf_df, inc_df)
            if not fcff_list:
                continue

            company_fcff_series[ts_code] = {f['year']: f['fcff'] for f in fcff_list}
            company_rev_series[ts_code] = self._extract_annual_revenue(inc_df)
            ni_map = self._extract_annual_net_income(inc_df)
            company_ni_series[ts_code] = ni_map

            # FCFF/Revenue ratios per year
            rev_map = company_rev_series[ts_code]
            for f in fcff_list:
                year = f['year']
                rev = rev_map.get(year)
                if rev and rev > 0:
                    ratio = f['fcff'] / rev
                    if -1 < ratio < 1:  # exclude extreme outliers
                        all_fcff_rev_ratios.append(ratio)

                # FCFF/NetIncome ratio
                ni = ni_map.get(year)
                if ni and ni > 0 and f['fcff'] > 0:
                    all_fcff_ni_ratios.append(f['fcff'] / ni)

            # FCFF YoY growth
            fcff_vals = [f['fcff'] for f in fcff_list]
            for i in range(1, len(fcff_vals)):
                if fcff_vals[i - 1] != 0:
                    g = (fcff_vals[i] - fcff_vals[i - 1]) / abs(fcff_vals[i - 1])
                    if -0.5 < g < 2.0:
                        all_fcff_growth_rates.append(g)

            # Revenue YoY growth
            rev_vals = sorted(rev_map.items())
            for i in range(1, len(rev_vals)):
                _, prev_rev = rev_vals[i - 1]
                _, curr_rev = rev_vals[i]
                if prev_rev > 0:
                    g = (curr_rev - prev_rev) / prev_rev
                    if -0.5 < g < 2.0:
                        all_rev_growth_rates.append(g)

        if not all_fcff_rev_ratios:
            return {'error': '行业无有效FCFF/营收数据'}

        # Statistics
        ratios = np.array(all_fcff_rev_ratios)
        fcff_g = np.array(all_fcff_growth_rates) if all_fcff_growth_rates else np.array([0.05])
        rev_g = np.array(all_rev_growth_rates) if all_rev_growth_rates else np.array([0.05])

        fcff_rev_median = float(np.median(ratios))
        fcff_growth_median = float(np.median(fcff_g))
        rev_growth_median = float(np.median(rev_g))
        maturity = self._determine_industry_maturity(rev_growth_median, len(companies))

        # Industry PE: from daily_basic real market data
        industry_pe = self._calc_industry_pe_from_market(industry_pe_data)

        # FCFF/NetIncome median ratio for PE → EV/FCFF conversion
        fcff_ni_median = float(np.median(all_fcff_ni_ratios)) if all_fcff_ni_ratios else 1.0

        # Build per-company regression dataset and fit OLS
        # Y = sustainable growth years (longest streak of FCFF growth > terminal_g)
        # X = [PE, avg_fcff_growth, avg_fcff_rev_ratio, avg_rev_growth]
        terminal_g = self._recommended_terminal_growth(maturity)
        pe_map = self._build_pe_map(industry_pe_data)

        # For companies without PE (NaN from Tushare), compute PE = market_cap / net_income
        pe_map = self._fill_missing_pe(pe_map, company_ni_series, industry_pe_data)

        regression_result = self._regression_forecast_years(
            company_fcff_series=company_fcff_series,
            company_rev_series=company_rev_series,
            company_ni_series=company_ni_series,
            pe_map=pe_map,
            terminal_g=terminal_g,
        )

        benchmark = {
            'l3_code': l3_code,
            'company_count': len(companies),
            'valid_company_count': len(company_fcff_series),

            'fcff_rev_ratio': {
                'median': fcff_rev_median,
                'mean': float(np.mean(ratios)),
                'p25': float(np.percentile(ratios, 25)),
                'p75': float(np.percentile(ratios, 75)),
                'std': float(np.std(ratios)),
            },

            'fcff_growth_rate': {
                'median': fcff_growth_median,
                'mean': float(np.mean(fcff_g)),
            },

            'revenue_growth_rate': {
                'median': rev_growth_median,
                'mean': float(np.mean(rev_g)),
            },

            'fcff_ni_ratio': round(fcff_ni_median, 4),

            'industry_pe': industry_pe,

            'industry_maturity': maturity,
            'recommended_forecast_years': regression_result['forecast_years'],
            'recommended_terminal_growth': terminal_g,
            'regression': regression_result,
        }

        return benchmark

    def _calc_company_fcff(
        self,
        cashflow_df: pd.DataFrame,
        income_df: pd.DataFrame = None,
        tax_rate: float = 0.25,
    ) -> list:
        """Calculate annual FCFF for a single company.

        FCFF = (OCF + Interest*(1-T) - CapEx) / 10000 (yuan → wan yuan)
        """
        if cashflow_df.empty:
            return []

        annual = cashflow_df[cashflow_df['end_date'].astype(str).str.contains('1231', na=False)].copy()
        annual = annual.sort_values('end_date')
        annual = annual.drop_duplicates(subset='end_date', keep='last')

        # Build interest lookup
        interest_by_year = {}
        if income_df is not None and not income_df.empty:
            inc_annual = income_df[income_df['end_date'].astype(str).str.contains('1231', na=False)].copy()
            inc_annual = inc_annual.drop_duplicates(subset='end_date', keep='last')
            for _, row in inc_annual.iterrows():
                year = str(row['end_date'])[:4]
                int_exp = row.get('fin_exp_int_exp', 0)
                if int_exp is None or (isinstance(int_exp, float) and np.isnan(int_exp)):
                    int_exp = 0
                interest_by_year[year] = float(int_exp)

        result = []
        for _, row in annual.iterrows():
            ocf = row.get('n_cashflow_act', 0)
            capex = row.get('c_pay_acq_const_fiolta', 0)
            year = str(row['end_date'])[:4]

            if isinstance(ocf, float) and np.isnan(ocf):
                ocf = 0
            if capex is None or (isinstance(capex, float) and np.isnan(capex)):
                capex = 0

            int_exp = interest_by_year.get(year, 0)
            after_tax_int = int_exp * (1 - tax_rate)

            if capex != 0:
                fcff = (ocf + after_tax_int - capex) / 10000
            else:
                fcff = (ocf + after_tax_int) * 0.7 / 10000

            result.append({'year': year, 'fcff': fcff})

        return result

    def _extract_annual_revenue(self, income_df: pd.DataFrame) -> dict:
        """Extract annual revenue (total_revenue) from income statement (wan yuan)."""
        if income_df is None or income_df.empty:
            return {}

        annual = income_df[income_df['end_date'].astype(str).str.contains('1231', na=False)].copy()
        annual = annual.drop_duplicates(subset='end_date', keep='last')

        result = {}
        for _, row in annual.iterrows():
            year = str(row['end_date'])[:4]
            rev = row.get('total_revenue', 0)
            if rev is None or (isinstance(rev, float) and np.isnan(rev)):
                continue
            result[year] = float(rev) / 10000  # yuan → wan yuan
        return result

    def _extract_annual_net_income(self, income_df: pd.DataFrame) -> dict:
        """Extract annual net income (n_income) from income statement (wan yuan)."""
        if income_df is None or income_df.empty:
            return {}

        annual = income_df[income_df['end_date'].astype(str).str.contains('1231', na=False)].copy()
        annual = annual.drop_duplicates(subset='end_date', keep='last')

        result = {}
        for _, row in annual.iterrows():
            year = str(row['end_date'])[:4]
            ni = row.get('n_income', 0)
            if ni is None or (isinstance(ni, float) and np.isnan(ni)):
                continue
            result[year] = float(ni) / 10000
        return result

    def _determine_industry_maturity(self, rev_growth_median: float, company_count: int) -> str:
        """Classify industry as 'growth', 'mature', or 'declining'."""
        if rev_growth_median > 0.15 and company_count > 10:
            return 'growth'
        if rev_growth_median > 0.03:
            return 'mature'
        return 'declining'

    def _recommended_terminal_growth(self, maturity: str) -> float:
        return {'growth': 0.03, 'mature': 0.025, 'declining': 0.02}.get(maturity, 0.025)

    def _calc_industry_pe_from_market(self, pe_data: pd.DataFrame) -> dict:
        """Calculate industry PE from daily_basic market data.

        Args:
            pe_data: DataFrame with 'pe' column from get_industry_daily_basics().

        Returns:
            PE statistics: median, mean, p25, p75, valid_count.
        """
        if pe_data is None or pe_data.empty or 'pe' not in pe_data.columns:
            return {'median': 0, 'mean': 0, 'p25': 0, 'p75': 0, 'valid_count': 0}

        # Filter valid PE: exclude NaN and extreme values, keep negative
        pe_vals = pe_data['pe'].dropna()
        pe_vals = pe_vals[pe_vals.abs() < 500]

        if pe_vals.empty:
            return {'median': 0, 'mean': 0, 'p25': 0, 'p75': 0, 'valid_count': 0}

        return {
            'median': round(float(pe_vals.median()), 1),
            'mean': round(float(pe_vals.mean()), 1),
            'p25': round(float(pe_vals.quantile(0.25)), 1),
            'p75': round(float(pe_vals.quantile(0.75)), 1),
            'valid_count': len(pe_vals),
        }

    def _build_pe_map(self, pe_data: pd.DataFrame) -> dict:
        """Build ts_code → PE mapping from daily_basic data."""
        if pe_data is None or pe_data.empty or 'pe' not in pe_data.columns:
            return {}
        valid = pe_data[pe_data['pe'].notna()]
        return dict(zip(valid['ts_code'], valid['pe']))

    def _fill_missing_pe(
        self,
        pe_map: dict,
        company_ni_series: dict,
        pe_data: pd.DataFrame,
    ) -> dict:
        """Fill PE for companies where Tushare reports NaN.

        Compute PE = total_mv / net_income using latest annual net income
        from financial statements and market cap from daily_basic.
        """
        if pe_data is None or pe_data.empty:
            return pe_map

        # Build market cap lookup
        mv_map = {}
        if 'total_mv' in pe_data.columns:
            for _, row in pe_data.iterrows():
                mv = row.get('total_mv')
                if mv and not (isinstance(mv, float) and np.isnan(mv)):
                    mv_map[row['ts_code']] = float(mv) * 10000  # 万元 → 元

        pe_map = dict(pe_map)  # copy

        for ts_code, ni_map in company_ni_series.items():
            if ts_code in pe_map:
                continue  # already has PE
            if not ni_map or ts_code not in mv_map:
                continue

            # Latest annual net income (wan yuan → yuan)
            latest_year = max(ni_map.keys())
            ni_wan = ni_map.get(latest_year, 0)
            if ni_wan == 0:
                continue
            ni_yuan = ni_wan * 10000

            market_cap = mv_map[ts_code]
            pe = market_cap / ni_yuan
            pe_map[ts_code] = round(pe, 1)

        return pe_map

    def _calc_sustainable_years(
        self,
        fcff_series: dict,
        terminal_g: float = 0.025,
    ) -> int:
        """Longest streak of consecutive years with FCFF growth > terminal_g."""
        if len(fcff_series) < 2:
            return 1

        sorted_years = sorted(fcff_series.keys())
        streak = 0
        max_streak = 0

        for i in range(1, len(sorted_years)):
            prev = fcff_series[sorted_years[i - 1]]
            curr = fcff_series[sorted_years[i]]
            if prev > 0 and curr > prev * (1 + terminal_g):
                streak += 1
                max_streak = max(max_streak, streak)
            else:
                streak = 0

        return max(1, max_streak)

    def _calc_realized_forecast_years(
        self,
        fcff_series: dict,
        from_year: str,
        forward_window: int = 5,
    ) -> int:
        """Realized forecast period: number of years with positive FCFF YoY growth
        in the next `forward_window` years after `from_year`.

        This is a forward-looking Y variable with good variance.
        Returns 0 if no future data available.
        """
        sorted_years = sorted(fcff_series.keys())
        if from_year not in sorted_years:
            return 0

        from_idx = sorted_years.index(from_year)
        future_years = sorted_years[from_idx + 1: from_idx + 1 + forward_window]

        if not future_years:
            return 0

        count = 0
        for fy in future_years:
            fy_idx = sorted_years.index(fy)
            prev_val = fcff_series[sorted_years[fy_idx - 1]]
            curr_val = fcff_series[fy]
            if prev_val != 0 and curr_val > prev_val:
                count += 1

        return count

    def _regression_forecast_years(
        self,
        company_fcff_series: dict,
        company_rev_series: dict,
        company_ni_series: dict,
        pe_map: dict,
        terminal_g: float = 0.025,
        default_wacc: float = 0.09,
    ) -> dict:
        """Derive forecast years via per-company reverse DCF, then OLS regression.

        Step 1: For each company with valid PE data:
            - Compute FCFF/NI ratio → target EV/FCFF = PE / (FCFF/NI)
            - Reverse DCF: find n where DCF(n, g, w, g_t) ≈ target multiple
            - High PE → more implied growth years needed

        Step 2: Fit OLS: Y = implied_n ~ PE + fcff_g + fcff_rev + rev_g

        Step 3: Per-year prediction using industry-median X values.
        """
        observations = []

        for ts_code, fcff_map in company_fcff_series.items():
            pe = pe_map.get(ts_code)
            if pe is None:
                continue  # no PE data at all
            if pe == 0:
                continue

            ni_map = company_ni_series.get(ts_code, {})
            rev_map = company_rev_series.get(ts_code, {})
            if not fcff_map or not ni_map:
                continue

            # Per-company FCFF/NI ratio (median of recent 3 years)
            fcff_ni_ratios = []
            sorted_years = sorted(fcff_map.keys(), reverse=True)
            count = 0
            for yr in sorted_years:
                ni = ni_map.get(yr)
                fcff = fcff_map.get(yr)
                if ni and ni > 0 and fcff and fcff > 0:
                    fcff_ni_ratios.append(fcff / ni)
                    count += 1
                    if count >= 3:
                        break

            if not fcff_ni_ratios:
                # Negative PE company with no valid FCFF/NI → implied_n = 1
                if pe < 0:
                    fcff_g = self._avg_growth(fcff_map)
                    fcff_rev = self._latest_ratio(fcff_map, rev_map)
                    rev_g = self._avg_growth(rev_map)
                    observations.append({
                        'ts_code': ts_code,
                        'implied_n': 1,
                        'pe': pe,
                        'fcff_g': fcff_g,
                        'fcff_rev': fcff_rev,
                        'rev_g': rev_g,
                    })
                continue
            fcff_ni = float(np.median(fcff_ni_ratios))
            if fcff_ni <= 0:
                continue

            fcff_g = self._avg_growth(fcff_map)
            fcff_rev = self._latest_ratio(fcff_map, rev_map)
            rev_g = self._avg_growth(rev_map)

            implied_n = self._reverse_dcf_implied_n(
                pe=pe,
                fcff_ni_ratio=fcff_ni,
                fcff_growth=fcff_g,
                wacc=default_wacc,
                terminal_g=terminal_g,
            )

            if implied_n is not None and implied_n > 0:
                observations.append({
                    'ts_code': ts_code,
                    'implied_n': implied_n,
                    'pe': pe,
                    'fcff_g': fcff_g,
                    'fcff_rev': fcff_rev,
                    'rev_g': rev_g,
                })

        if not observations:
            return self._default_regression_result()

        df_obs = pd.DataFrame(observations)
        n_obs = len(df_obs)
        df_obs['implied_n_clamped'] = df_obs['implied_n'].clip(3, 15)

        # --- OLS Regression ---
        X_features = ['pe', 'fcff_g', 'fcff_rev', 'rev_g']
        df_fit = df_obs.dropna(subset=X_features)
        n_fit = len(df_fit)

        if n_fit < max(10, len(X_features) + 3):
            median_n = int(round(df_obs['implied_n_clamped'].median()))
            return {
                'forecast_years': max(3, min(15, median_n)),
                'by_year': {},
                'r_squared': 0,
                'coefficients': {},
                'sample_size': n_obs,
                'method': f'反推DCF中位数 ({n_obs}家公司)',
            }

        Y = df_fit['implied_n_clamped'].values
        X_raw = df_fit[X_features].values
        X = np.column_stack([np.ones(n_fit), X_raw])

        try:
            beta = np.linalg.lstsq(X, Y, rcond=None)[0]
        except np.linalg.LinAlgError:
            median_n = int(round(df_obs['implied_n_clamped'].median()))
            return {
                'forecast_years': max(3, min(15, median_n)),
                'by_year': {},
                'r_squared': 0,
                'coefficients': {},
                'sample_size': n_obs,
                'method': f'反推DCF中位数 ({n_obs}家, 回归失败)',
            }

        Y_pred = X @ beta
        ss_res = np.sum((Y - Y_pred) ** 2)
        ss_tot = np.sum((Y - np.mean(Y)) ** 2)
        r_squared = 1 - ss_res / ss_tot if ss_tot > 0 else 0

        coef_names = ['intercept'] + X_features
        coefficients = {name: round(float(b), 4) for name, b in zip(coef_names, beta)}

        # --- Per-year prediction using industry-median X ---
        all_years = set()
        for fcff_map in company_fcff_series.values():
            all_years.update(fcff_map.keys())
        all_years = sorted(all_years)

        by_year = {}
        for year in all_years:
            year_pe_vals, year_fcff_gs, year_fcff_revs, year_rev_gs = [], [], [], []

            for ts_code, fcff_map in company_fcff_series.items():
                if year not in fcff_map:
                    continue
                rev_map = company_rev_series.get(ts_code, {})
                pe = pe_map.get(ts_code, 0)

                trunc_fcff = {yr: v for yr, v in fcff_map.items() if yr <= year}
                trunc_rev = {yr: v for yr, v in rev_map.items() if yr <= year}

                if pe > 0:
                    year_pe_vals.append(pe)
                year_fcff_gs.append(self._avg_growth(trunc_fcff))
                year_fcff_revs.append(self._latest_ratio(trunc_fcff, trunc_rev))
                year_rev_gs.append(self._avg_growth(trunc_rev))

            if not year_fcff_gs:
                continue

            medians = [
                float(np.median(year_pe_vals)) if year_pe_vals else float(df_obs['pe'].median()),
                float(np.median(year_fcff_gs)),
                float(np.median(year_fcff_revs)),
                float(np.median(year_rev_gs)),
            ]

            x_input = np.array([1.0] + medians)
            y_pred = float(x_input @ beta)
            by_year[year] = max(3, min(15, round(y_pred)))

        latest_year = max(by_year.keys()) if by_year else None
        forecast_years = by_year.get(latest_year, 5) if latest_year else 5

        return {
            'forecast_years': forecast_years,
            'by_year': by_year,
            'r_squared': round(r_squared, 4),
            'coefficients': coefficients,
            'sample_size': n_obs,
            'method': f'反推DCF+OLS回归 ({n_obs}家公司)',
        }

    def _reverse_dcf_implied_n(
        self,
        pe: float,
        fcff_ni_ratio: float,
        fcff_growth: float,
        wacc: float = 0.09,
        terminal_g: float = 0.025,
        max_years: int = 30,
    ) -> Optional[int]:
        """Reverse DCF: find implied forecast years n from PE and FCFF growth.

        Given a company's PE and FCFF/NI ratio, compute the target EV/FCFF multiple.
        Then find the smallest n where DCF(n, g, w, g_t) >= target.

        DCF_multiple(n) = sum_{t=1}^{n} (1+g)^t/(1+w)^t
                        + (1+g)^n * (1+g_t) / ((w-g_t) * (1+w)^n)

        High PE → high target multiple → more implied growth years.
        Negative PE → loss-making company → no growth years priced in → return 1.
        """
        if fcff_ni_ratio <= 0:
            return None
        if pe <= 0:
            return 1

        target = pe / fcff_ni_ratio  # EV/FCFF multiple implied by PE
        if target <= 1:
            return 1

        g = fcff_growth
        w = wacc
        g_t = terminal_g

        if w <= g_t:
            return None

        for n in range(1, max_years + 1):
            pv_fcff = sum((1 + g) ** t / (1 + w) ** t for t in range(1, n + 1))
            tv = (1 + g) ** n * (1 + g_t) / ((w - g_t) * (1 + w) ** n)
            multiple = pv_fcff + tv

            if multiple >= target:
                return n

        return max_years

    def _avg_growth(self, value_map: dict) -> float:
        """Average YoY growth rate from a {year: value} map."""
        if not value_map or len(value_map) < 2:
            return 0
        sorted_items = sorted(value_map.items())
        g_list = []
        for i in range(1, len(sorted_items)):
            prev = sorted_items[i - 1][1]
            curr = sorted_items[i][1]
            if prev != 0:
                g = (curr - prev) / abs(prev)
                if -0.5 < g < 2.0:
                    g_list.append(g)
        return float(np.mean(g_list)) if g_list else 0

    def _latest_ratio(self, fcff_map: dict, rev_map: dict) -> float:
        """Latest FCFF/Revenue ratio from two {year: value} maps."""
        for yr in sorted(fcff_map.keys(), reverse=True):
            rev = rev_map.get(yr)
            if rev and rev > 0 and fcff_map[yr]:
                r = fcff_map[yr] / rev
                if -1 < r < 1:
                    return r
        return 0

    def _default_regression_result(self) -> dict:
        return {
            'forecast_years': 5,
            'by_year': {},
            'r_squared': 0,
            'coefficients': {},
            'sample_size': 0,
            'method': '默认值 (数据不足)',
        }

    def _fallback_median_by_year(self, df_all: pd.DataFrame, n_fit: int) -> dict:
        by_year = {}
        for year in sorted(df_all['year'].unique()):
            df_yr = df_all[df_all['year'] == year]
            median_y = float(df_yr['y'].median())
            by_year[year] = max(3, min(15, round(median_y)))
        latest = max(by_year.keys()) if by_year else None
        return {
            'forecast_years': by_year.get(latest, 5) if latest else 5,
            'by_year': by_year,
            'r_squared': 0,
            'coefficients': {},
            'sample_size': n_fit,
            'method': f'median (样本{n_fit}<15, 不足回归)',
        }

    # ================================================================
    # Step 2: Individual Company Valuation
    # ================================================================

    def calculate_company_valuation(
        self,
        ts_code: str,
        industry_benchmark: dict,
        company_financials: dict,
        market_data: dict = None,
        params: dict = None,
    ) -> dict:
        """Perform industry-calibrated DCF valuation for a single company.

        Args:
            ts_code: Target stock code.
            industry_benchmark: Output from calculate_industry_benchmark().
            company_financials: {
                'cashflow': DataFrame or list-of-dicts,
                'income': DataFrame or list-of-dicts,
                'balance': DataFrame or list-of-dicts,
            }
            market_data: {
                'current_price': float,
                'total_shares': float (wan shares),
                'market_cap': float (yuan),
            }
            params: Override DEFAULT_PARAMS.

        Returns:
            Full valuation result dict.
        """
        if 'error' in industry_benchmark:
            return industry_benchmark

        p = dict(self.params)
        if params:
            p.update(params)

        # Normalize DataFrames
        cf_df = self._to_dataframe(company_financials.get('cashflow'))
        inc_df = self._to_dataframe(company_financials.get('income'))
        bs_df = self._to_dataframe(company_financials.get('balance'))

        # --- Company FCFF/Revenue ---
        company_fcff = self._calc_company_fcff(cf_df, inc_df, p['tax_rate'])
        company_rev = self._extract_annual_revenue(inc_df)

        if not company_fcff:
            return {'error': f'{ts_code} 无有效FCFF数据'}

        # Company's own FCFF/Revenue ratio (latest 3 years average)
        company_ratios = []
        for f in company_fcff[-3:]:
            rev = company_rev.get(f['year'])
            if rev and rev > 0:
                ratio = f['fcff'] / rev
                if -1 < ratio < 1:
                    company_ratios.append(ratio)

        company_ratio = float(np.mean(company_ratios)) if company_ratios else None

        # Industry benchmark ratio
        ind_ratio = industry_benchmark['fcff_rev_ratio']['median']

        # Adjustment coefficient
        alpha = self._calculate_alpha(company_ratio, ind_ratio)

        # Company revenue growth rate
        company_rev_growth = self._estimate_growth_from_series(company_rev)
        ind_growth = industry_benchmark['fcff_growth_rate']['median']
        blended_growth = self._blend_growth_rate(ind_growth, company_rev_growth, p['industry_weight'])

        # Latest revenue
        latest_year = max(company_rev.keys()) if company_rev else None
        latest_revenue = company_rev.get(latest_year, 0) if latest_year else 0

        if latest_revenue <= 0:
            return {'error': f'{ts_code} 营收数据无效'}

        # --- WACC ---
        market_data = market_data or {}
        wacc_info = self._calculate_wacc(
            ts_code, market_data, p, industry_benchmark,
        )

        # --- DCF Projection ---
        n_years = p.get('forecast_years', industry_benchmark['recommended_forecast_years'])
        terminal_g = p.get('terminal_growth_rate', industry_benchmark['recommended_terminal_growth'])
        wacc = wacc_info['wacc']

        projected = self._project_fcff_series(
            base_revenue=latest_revenue,
            industry_ratio=ind_ratio,
            alpha=alpha,
            growth_rate=blended_growth,
            n_years=n_years,
        )

        dcf = self._discount_fcff(projected, wacc, terminal_g, n_years)

        # Net debt
        net_debt = self._calculate_net_debt(bs_df)

        # Per-share value
        ev = dcf['enterprise_value']
        equity = ev - net_debt
        total_shares = market_data.get('total_shares', 0)
        current_price = market_data.get('current_price', 0)

        per_share = equity / total_shares if total_shares > 0 else 0
        margin = ((per_share / current_price) - 1) * 100 if current_price > 0 else 0

        # Sensitivity matrix
        sensitivity = self._generate_sensitivity(
            latest_revenue=latest_revenue,
            industry_ratio=ind_ratio,
            alpha=alpha,
            net_debt=net_debt,
            total_shares=total_shares,
            current_price=current_price,
            terminal_g=terminal_g,
            n_years=n_years,
        )

        return {
            'ts_code': ts_code,
            'valuation_date': pd.Timestamp.now().strftime('%Y-%m-%d'),

            # Company analysis
            'company_fcff_rev_ratio': round(company_ratio, 4) if company_ratio is not None else None,
            'industry_fcff_rev_ratio_median': round(ind_ratio, 4),
            'alpha': round(alpha, 4),
            'company_revenue_latest': round(latest_revenue, 2),
            'company_revenue_growth': round(company_rev_growth, 4),
            'company_fcff_history': company_fcff[-5:],

            # Growth parameters
            'industry_growth_rate': round(ind_growth, 4),
            'blended_growth_rate': round(blended_growth, 4),
            'terminal_growth_rate': terminal_g,
            'forecast_years': n_years,

            # WACC
            **wacc_info,

            # Projections
            'projected_fcff': projected,
            'pv_forecasts': round(dcf['pv_forecasts'], 2),
            'terminal_value': round(dcf['terminal_value'], 2),
            'pv_terminal': round(dcf['pv_terminal'], 2),
            'enterprise_value': round(ev, 2),
            'net_debt': round(net_debt, 2),
            'equity_value': round(equity, 2),

            # Per-share
            'total_shares': total_shares,
            'per_share_value': round(per_share, 2),
            'current_price': current_price,
            'safety_margin_pct': round(margin, 1),

            # Industry context
            'industry_benchmark': industry_benchmark,

            # Sensitivity
            'sensitivity': sensitivity,
        }

    def _calculate_alpha(self, company_ratio: Optional[float], industry_median: float) -> float:
        """Compute and clamp the adjustment coefficient."""
        if company_ratio is None or industry_median == 0:
            return 1.0
        alpha = company_ratio / industry_median
        return float(np.clip(alpha, self.ALPHA_MIN, self.ALPHA_MAX))

    def _blend_growth_rate(
        self,
        g_industry: float,
        g_company: float,
        weight_industry: float = 0.7,
    ) -> float:
        """Blend industry and company growth rates, clamped to [-10%, +20%]."""
        g = weight_industry * g_industry + (1 - weight_industry) * g_company
        return float(np.clip(g, -0.10, 0.20))

    def _estimate_growth_from_series(self, year_value_map: dict) -> float:
        """Estimate CAGR from a {year: value} map."""
        if not year_value_map or len(year_value_map) < 2:
            return 0.05
        sorted_items = sorted(year_value_map.items())
        values = [v for _, v in sorted_items]
        growth_rates = []
        for i in range(1, len(values)):
            if values[i - 1] != 0:
                g = (values[i] - values[i - 1]) / abs(values[i - 1])
                if -0.5 < g < 2.0:
                    growth_rates.append(g)
        if not growth_rates:
            return 0.05
        return float(np.clip(np.mean(growth_rates), -0.10, 0.20))

    def _calculate_wacc(
        self,
        ts_code: str,
        market_data: dict,
        params: dict,
        benchmark: dict,
    ) -> dict:
        """Calculate WACC using industry beta."""
        p = params
        beta = 1.0  # fallback

        if self.wacc_calc and self.wacc_calc.pro is not None:
            # Try industry beta via WACCCalculator
            try:
                beta_info = self.wacc_calc.calculate_beta_from_api(ts_code)
                beta = beta_info
            except Exception:
                beta = 1.0

        ke = p['risk_free_rate'] + beta * p['market_risk_premium']
        kd_aftertax = p['kd_pretax'] * (1 - p['tax_rate'])

        # Capital structure from market data
        market_cap = market_data.get('market_cap', 0)
        total_debt = market_data.get('total_debt', 0)
        if market_cap > 0 and (market_cap + total_debt) > 0:
            equity_ratio = market_cap / (market_cap + total_debt)
            debt_ratio = total_debt / (market_cap + total_debt)
        else:
            equity_ratio = 0.6
            debt_ratio = 0.4

        wacc = equity_ratio * ke + debt_ratio * kd_aftertax

        return {
            'wacc': round(wacc, 4),
            'ke': round(ke, 4),
            'kd_aftertax': round(kd_aftertax, 4),
            'beta': round(beta, 4),
            'equity_ratio': round(equity_ratio, 4),
            'debt_ratio': round(debt_ratio, 4),
        }

    def _project_fcff_series(
        self,
        base_revenue: float,
        industry_ratio: float,
        alpha: float,
        growth_rate: float,
        n_years: int,
    ) -> list:
        """Project future FCFF using industry-calibrated parameters."""
        projected = []
        rev = base_revenue
        for t in range(1, n_years + 1):
            rev = rev * (1 + growth_rate)
            fcff = rev * industry_ratio * alpha
            projected.append({
                'year': t,
                'revenue': round(rev, 2),
                'fcff': round(fcff, 2),
            })
        return projected

    def _discount_fcff(
        self,
        projected: list,
        wacc: float,
        terminal_g: float,
        n_years: int,
    ) -> dict:
        """Standard DCF: PV of forecasts + PV of terminal value (Gordon Growth)."""
        pv_forecasts = 0
        pv_list = []
        for item in projected:
            t = item['year']
            fcff = item['fcff']
            df = 1 / (1 + wacc) ** t
            pv = fcff * df
            pv_forecasts += pv
            pv_list.append({
                'year': t,
                'fcff': fcff,
                'discount_factor': round(df, 6),
                'pv': round(pv, 2),
            })

        # Update projected with PV info
        for i, item in enumerate(projected):
            item['pv'] = pv_list[i]['pv']

        # Terminal value
        last_fcff = projected[-1]['fcff'] if projected else 0
        terminal_fcff = last_fcff * (1 + terminal_g)
        if wacc > terminal_g:
            tv = terminal_fcff / (wacc - terminal_g)
            pv_tv = tv / (1 + wacc) ** n_years
        else:
            tv = 0
            pv_tv = 0

        return {
            'pv_forecasts': pv_forecasts,
            'terminal_value': tv,
            'pv_terminal': pv_tv,
            'enterprise_value': pv_forecasts + pv_tv,
        }

    def _calculate_net_debt(self, bs_df: pd.DataFrame) -> float:
        """Calculate net debt from balance sheet (wan yuan)."""
        if bs_df is None or bs_df.empty:
            return 0

        latest = bs_df.iloc[0]

        def _safe(row, key):
            v = row.get(key, 0)
            if v is None or (isinstance(v, float) and np.isnan(v)):
                return 0.0
            return float(v)

        st_borr = _safe(latest, 'st_borr')
        lt_borr = _safe(latest, 'lt_borr')
        bonds = _safe(latest, 'bond_payable')
        current_ltd = _safe(latest, 'non_cur_liab_due_1y')
        cash = _safe(latest, 'money_cap')

        total_debt = (st_borr + lt_borr + bonds + current_ltd) / 10000
        cash_equiv = cash / 10000
        return total_debt - cash_equiv

    def _generate_sensitivity(
        self,
        latest_revenue: float,
        industry_ratio: float,
        alpha: float,
        net_debt: float,
        total_shares: float,
        current_price: float,
        terminal_g: float,
        n_years: int,
    ) -> dict:
        """Generate lightweight sensitivity matrix for WACC × growth_rate."""
        if not total_shares or not current_price:
            return {}

        wacc_values = np.arange(0.07, 0.12 + 0.001, 0.01).tolist()
        growth_values = np.arange(0.02, 0.20 + 0.001, 0.04).tolist()

        matrix = []
        for w in wacc_values:
            row = []
            for g in growth_values:
                projected = []
                rev = latest_revenue
                for _ in range(n_years):
                    rev *= (1 + g)
                    projected.append(rev * industry_ratio * alpha)

                pv = sum(f / (1 + w) ** t for t, f in enumerate(projected, 1))
                last_f = projected[-1] if projected else 0
                tv = last_f * (1 + terminal_g) / (w - terminal_g) if w > terminal_g else 0
                pv_tv = tv / (1 + w) ** n_years
                ev = pv + pv_tv
                equity = ev - net_debt
                per_share = equity / total_shares
                margin = ((per_share / current_price) - 1) * 100
                row.append({'per_share': round(per_share, 2), 'margin_pct': round(margin, 1)})
            matrix.append(row)

        return {
            'wacc_values': [round(w, 4) for w in wacc_values],
            'growth_values': [round(g, 4) for g in growth_values],
            'matrix': matrix,
            'current_price': current_price,
        }

    @staticmethod
    def _to_dataframe(data) -> pd.DataFrame:
        """Convert list-of-dicts or DataFrame to DataFrame."""
        if data is None:
            return pd.DataFrame()
        if isinstance(data, pd.DataFrame):
            return data
        if isinstance(data, list):
            return pd.DataFrame(data)
        return pd.DataFrame()


# ================================================================
# Module-level convenience function for cross-module use
# ================================================================

_forecast_years_cache = {}  # {l3_code: (years, timestamp)}


def get_industry_forecast_years(
    ts_code: str,
    pro,
    force_refresh: bool = False,
) -> int:
    """Get industry-calibrated forecast years for a stock.

    Reuses industry_dcf/data/ JSON cache (24h validity).
    Returns 5 on any failure (safe fallback to original default).

    Args:
        ts_code: Stock code (e.g., '002001.SZ').
        pro: Tushare pro_api instance.
        force_refresh: Force re-fetch industry data.

    Returns:
        Recommended forecast years (int), typically 3-15.
    """
    import time as _time
    from .rate_limiter import RateLimiter as _RL
    from .shenwan_lookup import find_l3_by_code as _find_l3
    from .industry_data_fetcher import IndustryDataFetcher as _IDF

    if not ts_code or pro is None:
        return 5

    try:
        # 1. Find L3 industry
        industry_info = _find_l3(ts_code, pro)
        if not industry_info:
            return 5
        l3_code = industry_info['l3_code']

        # 2. Check memory cache (valid 24h)
        if not force_refresh and l3_code in _forecast_years_cache:
            cached_years, cached_ts = _forecast_years_cache[l3_code]
            if _time.time() - cached_ts < 86400:
                return cached_years

        # 3. Fetch data (reuses JSON file cache)
        rl = _RL()
        fetcher = _IDF(pro, rate_limiter=rl)
        calculator = IndustryDCFCalculator()

        industry_financials = fetcher.get_industry_financials(
            l3_code, force_refresh=force_refresh,
        )
        industry_pe_data = fetcher.get_industry_daily_basics(l3_code)

        # 4. Calculate benchmark (reverse DCF + regression)
        benchmark = calculator.calculate_industry_benchmark(
            industry_financials, industry_pe_data=industry_pe_data,
        )

        if 'error' in benchmark:
            return 5

        years = benchmark.get('recommended_forecast_years', 5)

        # 5. Cache result
        _forecast_years_cache[l3_code] = (years, _time.time())

        return years

    except Exception:
        return 5
