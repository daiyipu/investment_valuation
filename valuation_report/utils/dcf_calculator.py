#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Unified DCF Calculator for A-Share Valuation.

Three methods: FCFF (WACC-based), FCFE (direct equity), APV (Adjusted Present Value).
Terminal value via Gordon Growth Model. Growth rate clamped to [-10%, +20%].

Usage:
    from dcf_calculator import DCFCalculator

    calc = DCFCalculator()
    result = calc.calculate_dcf_valuation(
        financial_statements={'cashflow': cf_df, 'balancesheet': bs_df, 'income': inc_df},
        params={'risk_free_rate': 0.0185, 'forecast_years': 10},
    )
    summary = calc.generate_valuation_summary(result, current_price=25.0, total_shares=100000)
"""

import numpy as np
import pandas as pd
from typing import Optional, Dict, Any, List


class DCFCalculator:
    """DCF valuation with three methods and sensitivity analysis."""

    def __init__(self):
        self.fcf_data = []

    def _prepare_fcf_data(self, cashflow_df: pd.DataFrame) -> list:
        """Extract historical FCF from cashflow statement.

        Priority:
        1. FCF = OCF - CapEx (n_cashflow_act - c_pay_acq_const_fiolta)
        2. FCF = OCF + Investing CF (n_cashflow_act + n_cashflow_inv_act)
        3. FCF = OCF × 0.7

        Tushare data unit: yuan → converted to wan yuan (÷10000).
        """
        if cashflow_df is None or cashflow_df.empty:
            return []

        annual = cashflow_df[cashflow_df['end_date'].astype(str).str.contains('1231', na=False)].copy()
        annual = annual.sort_values('end_date')
        annual = annual.drop_duplicates(subset='end_date', keep='last')

        fcf_list = []
        for _, row in annual.iterrows():
            ocf = row.get('n_cashflow_act', 0)
            capex = row.get('c_pay_acq_const_fiolta', 0)
            inv_cf = row.get('n_cashflow_inv_act', 0)
            # Handle NaN: convert to 0
            if isinstance(ocf, float) and np.isnan(ocf):
                ocf = 0
            if capex is None or (isinstance(capex, float) and np.isnan(capex)):
                capex = 0
            if inv_cf is None or (isinstance(inv_cf, float) and np.isnan(inv_cf)):
                inv_cf = 0

            if capex != 0:
                fcf = (ocf - capex) / 10000
            elif inv_cf != 0:
                fcf = (ocf + inv_cf) / 10000
            else:
                fcf = ocf * 0.7 / 10000

            fcf_list.append({
                'year': str(row['end_date'])[:4],
                'fcf': fcf,
                'method': 'OCF-CapEx' if capex != 0 else ('OCF+InvCF' if inv_cf != 0 else 'OCF×0.7'),
            })

        self.fcf_data = fcf_list
        return fcf_list

    def _estimate_growth_rate(self, fcf_list: list, clamp: tuple = (-0.10, 0.20)) -> float:
        """Estimate FCF growth rate from historical data, clamped."""
        if len(fcf_list) < 2:
            return 0.05

        fcfs = [f['fcf'] for f in fcf_list if f['fcf'] != 0]
        if len(fcfs) < 2:
            return 0.05

        growth_rates = []
        for i in range(1, len(fcfs)):
            if fcfs[i - 1] != 0:
                g = (fcfs[i] - fcfs[i - 1]) / abs(fcfs[i - 1])
                if -0.5 < g < 2.0:
                    growth_rates.append(g)

        if not growth_rates:
            return 0.05

        avg_growth = np.mean(growth_rates)
        return float(np.clip(avg_growth, clamp[0], clamp[1]))

    def _calculate_net_debt(self, balancesheet_df: pd.DataFrame) -> float:
        """Calculate net debt from balance sheet (wan yuan)."""
        if balancesheet_df is None or balancesheet_df.empty:
            return 0

        latest = balancesheet_df.iloc[0]

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

    def _get_debt_ratio(self, balancesheet_df: pd.DataFrame) -> float:
        """Get debt-to-assets ratio from balance sheet."""
        if balancesheet_df is None or balancesheet_df.empty:
            return 0.4
        latest = balancesheet_df.iloc[0]
        total_liab = latest.get('total_liab', 0)
        total_assets = latest.get('total_assets', 0)
        if total_liab is None or total_assets is None:
            return 0.4
        total_liab = float(total_liab)
        total_assets = float(total_assets)
        if np.isnan(total_liab) or np.isnan(total_assets) or total_assets <= 0:
            return 0.4
        return total_liab / total_assets

    def calculate_dcf_valuation(
        self,
        financial_statements: dict,
        financial_indicators: pd.DataFrame = None,
        params: Optional[dict] = None,
    ) -> dict:
        """Run full three-method DCF valuation.

        Args:
            financial_statements: {
                'cashflow': DataFrame,
                'balancesheet': DataFrame,
                'income': DataFrame (optional),
            }
            financial_indicators: fina_indicator DataFrame (optional)
            params: override default calculation parameters

        Returns:
            Dict with WACC, three method results, and input parameters.
        """
        p = {
            'risk_free_rate': 0.0185,
            'market_risk_premium': 0.0615,
            'beta': 1.0,
            'kd': 0.05,
            'tax_rate': 0.25,
            'terminal_growth_rate': 0.025,
            'forecast_years': 10,
        }
        if params:
            p.update(params)

        cashflow_df = financial_statements.get('cashflow')
        bs_df = financial_statements.get('balancesheet')
        income_df = financial_statements.get('income')

        # 1. Prepare FCF data
        fcf_list = self._prepare_fcf_data(cashflow_df)
        if not fcf_list:
            return {'error': 'No FCF data available'}

        # 2. Estimate growth rate
        growth_rate = self._estimate_growth_rate(fcf_list)

        # 3. Capital structure
        debt_ratio = self._get_debt_ratio(bs_df)
        equity_ratio = 1 - debt_ratio

        # 4. WACC calculation
        ke = p['risk_free_rate'] + p['beta'] * p['market_risk_premium']
        kd_aftertax = p['kd'] * (1 - p['tax_rate'])
        wacc = ke * equity_ratio + kd_aftertax * debt_ratio

        # 5. Net debt
        net_debt = self._calculate_net_debt(bs_df)

        # 6. Project FCF
        last_fcf = fcf_list[-1]['fcf']
        n_years = p['forecast_years']
        terminal_g = p['terminal_growth_rate']

        projected_fcfs = []
        current_fcf = last_fcf
        for i in range(1, n_years + 1):
            current_fcf = current_fcf * (1 + growth_rate)
            projected_fcfs.append(current_fcf)

        # ---- Method 1: FCFF (Enterprise Value) ----
        pv_forecasts_1 = sum(fcf / (1 + wacc) ** t for t, fcf in enumerate(projected_fcfs, 1))
        terminal_fcf_1 = projected_fcfs[-1] * (1 + terminal_g)
        tv_1 = terminal_fcf_1 / (wacc - terminal_g) if wacc > terminal_g else 0
        pv_tv_1 = tv_1 / (1 + wacc) ** n_years
        ev_1 = pv_forecasts_1 + pv_tv_1

        cf_list_1 = []
        for i, fcf in enumerate(projected_fcfs):
            df = 1 / (1 + wacc) ** (i + 1)
            cf_list_1.append({
                'year': i + 1,
                'fcf': fcf,
                'discount_factor': df,
                'pv': fcf * df,
            })

        # ---- Method 2: FCFE (Direct Equity) ----
        fcfe_list = [fcf * equity_ratio for fcf in projected_fcfs]
        pv_forecasts_2 = sum(fcfe / (1 + ke) ** t for t, fcfe in enumerate(fcfe_list, 1))
        terminal_fcfe = fcfe_list[-1] * (1 + terminal_g)
        tv_2 = terminal_fcfe / (ke - terminal_g) if ke > terminal_g else 0
        pv_tv_2 = tv_2 / (1 + ke) ** n_years
        equity_2 = pv_forecasts_2 + pv_tv_2

        cf_list_2 = []
        for i, fcfe in enumerate(fcfe_list):
            df = 1 / (1 + ke) ** (i + 1)
            cf_list_2.append({
                'year': i + 1,
                'fcfe': fcfe,
                'discount_factor': df,
                'pv': fcfe * df,
            })

        # ---- Method 3: APV (Adjusted Present Value) ----
        # Unlevered beta: Bu = Bl / (1 + (1-t) * D/E)
        d_over_e = debt_ratio / equity_ratio if equity_ratio > 0 else 0
        unlevered_beta = p['beta'] / (1 + (1 - p['tax_rate']) * d_over_e)
        unlevered_beta = float(np.clip(unlevered_beta, 0.2, 2.5))
        ku = p['risk_free_rate'] + unlevered_beta * p['market_risk_premium']

        # Unlevered firm value: PV of FCFF discounted at Ku
        pv_forecasts_3 = sum(fcf / (1 + ku) ** t for t, fcf in enumerate(projected_fcfs, 1))
        terminal_fcf_3 = projected_fcfs[-1] * (1 + terminal_g)
        tv_3 = terminal_fcf_3 / (ku - terminal_g) if ku > terminal_g else 0
        pv_tv_3 = tv_3 / (1 + ku) ** n_years
        unlevered_value = pv_forecasts_3 + pv_tv_3

        # PV of tax shield: sum of (tax_rate * Kd * debt_i) discounted
        # Assume debt grows with FCF, so debt_i scales with projected FCF
        total_debt_wan = net_debt + (bs_df.iloc[0].get('money_cap', 0) or 0) / 10000 if bs_df is not None and not bs_df.empty else net_debt
        if total_debt_wan < 0:
            total_debt_wan = 0
        pv_tax_shield = 0
        debt_base = total_debt_wan * debt_ratio if total_debt_wan > 0 else last_fcf * debt_ratio / max(wacc, 0.01)
        for t in range(1, n_years + 1):
            tax_shield_t = p['tax_rate'] * p['kd'] * debt_base
            pv_tax_shield += tax_shield_t / (1 + p['kd']) ** t
            debt_base *= (1 + growth_rate)

        apv_firm_value = unlevered_value + pv_tax_shield
        equity_3 = apv_firm_value - net_debt

        cf_list_3 = []
        for i, fcf in enumerate(projected_fcfs):
            df = 1 / (1 + ku) ** (i + 1)
            cf_list_3.append({
                'year': i + 1,
                'fcf': fcf,
                'discount_factor': df,
                'pv': fcf * df,
            })

        return {
            # Input parameters
            'wacc': wacc,
            'ke': ke,
            'kd_aftertax': kd_aftertax,
            'growth_rate': growth_rate,
            'terminal_growth_rate': terminal_g,
            'equity_ratio': equity_ratio,
            'debt_ratio': debt_ratio,
            'forecast_years': n_years,
            'params': p,

            # Historical data
            'historical_fcf': fcf_list,
            'last_fcf': last_fcf,

            # Method 1: FCFF
            'method1_ev': ev_1,
            'method1_pv_forecasts': pv_forecasts_1,
            'method1_terminal_value': tv_1,
            'method1_pv_terminal': pv_tv_1,
            'method1_cf_list': cf_list_1,

            # Method 2: FCFE
            'method2_equity_value': equity_2,
            'method2_pv_forecasts': pv_forecasts_2,
            'method2_terminal_value': tv_2,
            'method2_pv_terminal': pv_tv_2,
            'method2_cf_list': cf_list_2,

            # Method 3: APV
            'method3_equity_value': equity_3,
            'method3_apv_firm_value': apv_firm_value,
            'method3_unlevered_value': unlevered_value,
            'method3_pv_tax_shield': pv_tax_shield,
            'method3_pv_forecasts': pv_forecasts_3,
            'method3_terminal_value': tv_3,
            'method3_pv_terminal': pv_tv_3,
            'method3_unlevered_beta': unlevered_beta,
            'method3_unlevered_ke': ku,
            'method3_cf_list': cf_list_3,
            'net_debt': net_debt,

            # Capital structure
            'market_cap': 0,  # needs external data
        }

    def generate_valuation_summary(
        self,
        dcf_result: dict,
        current_price: float = 0,
        total_shares: float = 0,
    ) -> pd.DataFrame:
        """Generate per-share value summary with safety margins.

        Args:
            dcf_result: output from calculate_dcf_valuation()
            current_price: latest stock price
            total_shares: total shares outstanding (wan shares)

        Returns:
            DataFrame with per-share values and margins for each method.
        """
        if not total_shares or total_shares <= 0:
            return pd.DataFrame()

        rows = []

        # Method 1: FCFF → Equity = EV - Net Debt
        if dcf_result.get('method1_ev'):
            ev = dcf_result['method1_ev']
            net_debt = dcf_result.get('net_debt', 0)
            equity = ev - net_debt
            per_share = equity / total_shares
            margin = ((per_share / current_price) - 1) * 100 if current_price > 0 else 0
            rows.append({
                'method': 'FCFF (EV-NetDebt)',
                'enterprise_value': ev,
                'equity_value': equity,
                'per_share': per_share,
                'margin_pct': margin,
            })

        # Method 2: FCFE
        if dcf_result.get('method2_equity_value'):
            equity = dcf_result['method2_equity_value']
            per_share = equity / total_shares
            margin = ((per_share / current_price) - 1) * 100 if current_price > 0 else 0
            rows.append({
                'method': 'FCFE',
                'enterprise_value': '',
                'equity_value': equity,
                'per_share': per_share,
                'margin_pct': margin,
            })

        # Method 3: APV
        if dcf_result.get('method3_equity_value'):
            equity = dcf_result['method3_equity_value']
            per_share = equity / total_shares
            margin = ((per_share / current_price) - 1) * 100 if current_price > 0 else 0
            rows.append({
                'method': 'APV (调整现值法)',
                'enterprise_value': dcf_result.get('method3_apv_firm_value', ''),
                'equity_value': equity,
                'per_share': per_share,
                'margin_pct': margin,
            })

        return pd.DataFrame(rows)

    def generate_sensitivity_matrix(
        self,
        dcf_result: dict,
        current_price: float,
        total_shares: float,
        wacc_range: tuple = (0.07, 0.12, 0.01),
        growth_range: tuple = (0.02, 0.20, 0.02),
    ) -> dict:
        """Generate sensitivity matrix for per-share value.

        Args:
            dcf_result: output from calculate_dcf_valuation()
            current_price: latest stock price
            total_shares: total shares outstanding (wan shares)
            wacc_range: (start, end, step) for WACC axis
            growth_range: (start, end, step) for growth rate axis

        Returns:
            {
                'wacc_values': [...],
                'growth_values': [...],
                'matrix': [[per_share_value, margin_pct], ...],
                'current_price': price,
            }
        """
        if not total_shares or not current_price:
            return {}

        last_fcf = dcf_result.get('last_fcf', 0)
        net_debt = dcf_result.get('net_debt', 0)
        terminal_g = dcf_result.get('terminal_growth_rate', 0.025)
        n_years = dcf_result.get('forecast_years', 10)

        wacc_values = np.arange(wacc_range[0], wacc_range[1] + 0.001, wacc_range[2]).tolist()
        growth_values = np.arange(growth_range[0], growth_range[1] + 0.001, growth_range[2]).tolist()

        matrix = []
        for w in wacc_values:
            row = []
            for g in growth_values:
                # Project FCF with this growth rate
                projected = []
                fcf = last_fcf
                for _ in range(n_years):
                    fcf = fcf * (1 + g)
                    projected.append(fcf)

                # DCF
                pv = sum(f / (1 + w) ** t for t, f in enumerate(projected, 1))
                terminal_fcf = projected[-1] * (1 + terminal_g)
                tv = terminal_fcf / (w - terminal_g) if w > terminal_g else 0
                pv_tv = tv / (1 + w) ** n_years
                ev = pv + pv_tv
                equity = ev - net_debt
                per_share = equity / total_shares
                margin = ((per_share / current_price) - 1) * 100

                row.append({
                    'per_share': round(per_share, 2),
                    'margin_pct': round(margin, 1),
                })
            matrix.append(row)

        return {
            'wacc_values': [round(w, 4) for w in wacc_values],
            'growth_values': [round(g, 4) for g in growth_values],
            'matrix': matrix,
            'current_price': current_price,
        }
