#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Unified WACC Calculator for A-Share DCF Valuation.

Supports:
- API-computed Beta (stock returns vs CSI 300)
- Industry Beta (median of SW L3 peer Betas)
- Market-value based capital structure
- CAPM cost of equity
- Configurable parameters

Usage:
    from wacc_calculator import WACCCalculator

    calc = WACCCalculator(pro=pro)
    result = calc.calculate_wacc(stock_code='002276.SZ')
    # result = {wacc, ke, kd_pretax, kd_aftertax, beta, industry_beta, ...}

    # Or standalone (no API):
    result = calculate_wacc_simple(beta=1.2, equity_ratio=0.7)
"""

import numpy as np
from typing import Optional, Dict, Any


# Default parameters (2025 year-end)
DEFAULT_PARAMS = {
    'risk_free_rate': 0.0185,       # 10yr CGB yield
    'market_return': 0.08,          # CSI 300 expected return
    'market_risk_premium': 0.0615,  # Rm - Rf
    'debt_premium': 0.50,           # Kd = Rf × (1 + premium)
    'tax_rate': 0.25,
    'beta_window': 500,             # trading days (~2 years)
    'fallback_beta': 1.0,
}


class WACCCalculator:
    """WACC calculator with Tushare API integration."""

    def __init__(self, pro=None, params: Optional[dict] = None):
        self.pro = pro
        self.params = {**DEFAULT_PARAMS, **(params or {})}

    # ---- Beta Calculation ----

    def calculate_beta(self, stock_returns: np.ndarray,
                       market_returns: np.ndarray) -> float:
        """Compute Beta from return arrays via covariance/variance."""
        if len(stock_returns) < 30 or len(market_returns) < 30:
            return self.params['fallback_beta']
        min_len = min(len(stock_returns), len(market_returns))
        sr = stock_returns[-min_len:]
        mr = market_returns[-min_len:]
        cov_matrix = np.cov(sr, mr)
        beta = cov_matrix[0, 1] / np.var(mr)
        return np.clip(beta, 0.2, 2.5)

    def calculate_beta_from_api(self, stock_code: str,
                                 window: int = 0) -> float:
        """Compute Beta from Tushare daily data vs CSI 300."""
        if self.pro is None:
            return self.params['fallback_beta']

        window = window or self.params['beta_window']
        from datetime import datetime, timedelta
        end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.now() - timedelta(days=int(window * 1.8))).strftime('%Y%m%d')

        try:
            stock = self.pro.daily(
                ts_code=stock_code, start_date=start_date, end_date=end_date,
                fields='trade_date,pct_chg')
            market = self.pro.index_daily(
                ts_code='000300.SH', start_date=start_date, end_date=end_date,
                fields='trade_date,pct_chg')

            if stock is None or stock.empty or market is None or market.empty:
                return self.params['fallback_beta']

            stock = stock.sort_values('trade_date')
            market = market.sort_values('trade_date')

            merged = stock.merge(market, on='trade_date', suffixes=('_s', '_m'))
            if len(merged) < 60:
                return self.params['fallback_beta']

            sr = merged['pct_chg_s'].values / 100
            mr = merged['pct_chg_m'].values / 100
            return self.calculate_beta(sr, mr)

        except Exception as e:
            print(f"  Beta calculation failed: {e}")
            return self.params['fallback_beta']

    def calculate_industry_beta(self, stock_code: str,
                                 peer_codes: list = None,
                                 l3_code: str = None) -> float:
        """Compute industry Beta as median of peer Betas."""
        if self.pro is None:
            return self.params['fallback_beta']

        codes = peer_codes or []
        if not codes and l3_code:
            try:
                members = self.pro.index_member(index_code=l3_code)
                if members is not None and not members.empty:
                    codes = members['con_code'].tolist()
            except Exception:
                pass

        if not codes:
            return self.params['fallback_beta']

        betas = []
        for code in codes[:20]:  # cap at 20 peers
            try:
                b = self.calculate_beta_from_api(code)
                betas.append(b)
            except Exception:
                continue

        if not betas:
            return self.params['fallback_beta']

        return float(np.median(betas))

    # ---- Capital Structure ----

    def get_capital_structure(self, stock_code: str,
                               trade_date: str = '',
                               market_cap_yuan: float = 0) -> dict:
        """Fetch market-value based capital structure from Tushare."""
        from datetime import datetime, timedelta

        if not trade_date:
            trade_date = datetime.now().strftime('%Y%m%d')

        result = {
            'market_cap': 0,
            'total_debt': 0,
            'equity_ratio': 0.6,
            'debt_ratio': 0.4,
            'debt_breakdown': {},
        }

        if self.pro is None:
            return result

        try:
            # Market cap: 先尝试指定日期，失败则往前推
            market_cap = market_cap_yuan
            if market_cap <= 0:
                for days_back in range(0, 8):
                    td = (datetime.now() - timedelta(days=days_back)).strftime('%Y%m%d')
                    db = self.pro.daily_basic(
                        ts_code=stock_code, trade_date=td,
                        fields='trade_date,total_mv,total_share')
                    if db is not None and not db.empty:
                        market_cap = float(db.iloc[0]['total_mv']) * 10000
                        break
            result['market_cap'] = market_cap

            # Interest-bearing debt from balance sheet
            bs = self.pro.balancesheet(
                ts_code=stock_code,
                fields='ts_code,end_date,st_borr,lt_borr,bond_payable,'
                       'non_cur_liab_due_1y,total_hldr_eqy_inc_min_int')
            if bs is not None and not bs.empty:
                latest = bs.iloc[0]
                st_borr = float(latest.get('st_borr', 0) or 0)
                lt_borr = float(latest.get('lt_borr', 0) or 0)
                bonds = float(latest.get('bond_payable', 0) or 0)
                current_ltd = float(latest.get('non_cur_liab_due_1y', 0) or 0)
                total_debt = st_borr + lt_borr + bonds + current_ltd

                result['total_debt'] = total_debt
                result['debt_breakdown'] = {
                    'short_term_borrowing': st_borr,
                    'long_term_borrowing': lt_borr,
                    'bonds_payable': bonds,
                    'current_portion_ltd': current_ltd,
                }

                total_capital = result['market_cap'] + total_debt
                if total_capital > 0:
                    result['equity_ratio'] = result['market_cap'] / total_capital
                    result['debt_ratio'] = total_debt / total_capital

        except Exception as e:
            print(f"  Capital structure fetch failed: {e}")

        return result

    # ---- Full WACC ----

    def calculate_wacc(self, stock_code: str,
                       peer_codes: list = None,
                       use_industry_beta: bool = True) -> dict:
        """Full WACC calculation with API data."""
        p = self.params

        # Beta
        stock_beta = self.calculate_beta_from_api(stock_code)
        industry_beta = stock_beta
        if use_industry_beta and peer_codes:
            industry_beta = self.calculate_industry_beta(
                stock_code, peer_codes=peer_codes)

        adopted_beta = industry_beta if use_industry_beta else stock_beta

        # Capital structure
        cap_struct = self.get_capital_structure(stock_code)

        # Cost of Equity (CAPM)
        ke = p['risk_free_rate'] + adopted_beta * p['market_risk_premium']

        # Cost of Debt
        kd_pretax = p['risk_free_rate'] * (1 + p['debt_premium'])
        kd_aftertax = kd_pretax * (1 - p['tax_rate'])

        # WACC
        wacc = (cap_struct['equity_ratio'] * ke
                + cap_struct['debt_ratio'] * kd_aftertax)

        return {
            'wacc': wacc,
            'ke': ke,
            'kd_pretax': kd_pretax,
            'kd_aftertax': kd_aftertax,
            'stock_beta': stock_beta,
            'industry_beta': industry_beta,
            'adopted_beta': adopted_beta,
            'risk_free_rate': p['risk_free_rate'],
            'market_risk_premium': p['market_risk_premium'],
            'tax_rate': p['tax_rate'],
            'equity_ratio': cap_struct['equity_ratio'],
            'debt_ratio': cap_struct['debt_ratio'],
            'market_cap': cap_struct['market_cap'],
            'total_debt': cap_struct['total_debt'],
            'debt_breakdown': cap_struct['debt_breakdown'],
        }


def calculate_wacc_simple(
    beta: float = 1.0,
    risk_free_rate: float = 0.0185,
    market_risk_premium: float = 0.0615,
    equity_ratio: float = 0.6,
    debt_ratio: float = 0.4,
    kd_pretax: float = 0.0278,
    tax_rate: float = 0.25,
) -> dict:
    """Standalone WACC without API dependency."""
    ke = risk_free_rate + beta * market_risk_premium
    kd_aftertax = kd_pretax * (1 - tax_rate)
    wacc = equity_ratio * ke + debt_ratio * kd_aftertax

    return {
        'wacc': wacc,
        'ke': ke,
        'kd_pretax': kd_pretax,
        'kd_aftertax': kd_aftertax,
        'beta': beta,
        'equity_ratio': equity_ratio,
        'debt_ratio': debt_ratio,
    }
