#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DCF估值计算工具
三种估值方法：
1. 企业价值(EV) = Σ FCFF/(1+WACC)^t + TV/(1+WACC)^n
2. 股权价值 = Σ FCFE/(1+Ke)^t + TV/(1+Ke)^n
3. 股权价值(间接) = EV - 净债务
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Optional, Tuple, List


class DCFCalculator:
    """DCF估值计算器"""

    def __init__(self):
        """初始化DCF计算器"""
        pass

    def calculate_dcf_valuation(
        self,
        financial_statements: Dict[str, pd.DataFrame],
        financial_indicators: pd.DataFrame,
        risk_free_rate: float = 0.03,
        market_risk_premium: float = 0.06,
        beta: float = 1.0,
        terminal_growth_rate: float = 0.03,
        forecast_years: int = 5
    ) -> Dict[str, Any]:
        """计算DCF估值

        Args:
            financial_statements: 财务报表数据
            financial_indicators: 财务指标数据
            risk_free_rate: 无风险利率
            market_risk_premium: 市场风险溢价
            beta: Beta系数
            terminal_growth_rate: 永续增长率
            forecast_years: 预测年数

        Returns:
            DCF估值结果
        """
        result = {}

        # 从资产负债表计算实际债务比例
        debt_ratio = self._get_debt_ratio(financial_statements)

        # 计算折现率
        wacc = self._calculate_wacc(risk_free_rate, market_risk_premium, beta, debt_ratio)
        ke = self._calculate_ke(risk_free_rate, market_risk_premium, beta)

        result['wacc'] = wacc
        result['ke'] = ke
        result['risk_free_rate'] = risk_free_rate
        result['market_risk_premium'] = market_risk_premium
        result['beta'] = beta
        result['terminal_growth_rate'] = terminal_growth_rate
        result['debt_ratio'] = debt_ratio

        # 获取历史FCF数据并预测
        fcf_data = self._prepare_fcf_data(financial_statements, financial_indicators)

        if not fcf_data.empty:
            # 方法1: FCFF法计算企业价值
            ev, terminal_value_ff, cf_list_ff = self._calculate_ev_from_fcff(
                fcf_data, wacc, terminal_growth_rate, forecast_years
            )
            result['method1_ev'] = ev
            result['method1_terminal_value'] = terminal_value_ff
            result['method1_cf_list'] = cf_list_ff

            # 方法2: FCFE法计算股权价值（使用实际债务比例）
            equity_value_fe, terminal_value_fe, cf_list_fe = self._calculate_equity_from_fcfe(
                fcf_data, ke, terminal_growth_rate, forecast_years, debt_ratio
            )
            result['method2_equity_value'] = equity_value_fe
            result['method2_terminal_value'] = terminal_value_fe
            result['method2_cf_list'] = cf_list_fe

            # 方法3: 间接法计算股权价值
            net_debt = self._calculate_net_debt(financial_statements)
            equity_value_indirect = ev - net_debt
            result['method3_equity_value'] = equity_value_indirect
            result['net_debt'] = net_debt

        return result

    def _get_debt_ratio(self, financial_statements: Dict[str, pd.DataFrame]) -> float:
        """从资产负债表计算实际资产负债率

        Args:
            financial_statements: 财务报表

        Returns:
            债务比例(D/A)
        """
        balance_sheet = financial_statements.get('balance_sheet', pd.DataFrame())
        if balance_sheet.empty:
            return 0.4

        # 获取最新一期数据
        if 'end_date' in balance_sheet.columns:
            bs_sorted = balance_sheet.sort_values('end_date', ascending=False)
            latest = bs_sorted.iloc[0]
        else:
            latest = balance_sheet.iloc[-1]

        total_assets = latest.get('total_assets')
        total_liab = latest.get('total_liab')

        if (total_assets is not None and pd.notna(total_assets) and total_assets > 0
                and total_liab is not None and pd.notna(total_liab)):
            return float(total_liab) / float(total_assets)

        return 0.4

    def _calculate_wacc(
        self,
        risk_free_rate: float,
        market_risk_premium: float,
        beta: float,
        debt_ratio: float
    ) -> float:
        """计算WACC

        Args:
            risk_free_rate: 无风险利率
            market_risk_premium: 市场风险溢价
            beta: Beta系数
            debt_ratio: 债务比例(D/A)

        Returns:
            WACC
        """
        ke = risk_free_rate + beta * market_risk_premium
        kd = 0.05
        equity_ratio = 1 - debt_ratio
        tax_rate = 0.25
        wacc = ke * equity_ratio + kd * (1 - tax_rate) * debt_ratio
        return wacc

    def _calculate_ke(
        self,
        risk_free_rate: float,
        market_risk_premium: float,
        beta: float
    ) -> float:
        """计算权益成本Ke"""
        return risk_free_rate + beta * market_risk_premium

    def _prepare_fcf_data(
        self,
        financial_statements: Dict[str, pd.DataFrame],
        financial_indicators: pd.DataFrame
    ) -> pd.DataFrame:
        """准备FCF数据

        使用Tushare标准英文字段名获取现金流数据
        FCFF ≈ 经营活动现金流净额 + 投资活动现金流净额
        """
        cashflow = financial_statements.get('cashflow', pd.DataFrame())

        if cashflow.empty:
            return pd.DataFrame()

        cashflow = cashflow.copy()

        # 查找经营和投资活动现金流净额列（优先Tushare英文字段名）
        oper_col = self._find_column(
            cashflow,
            ['n_cashflow_act'],
            ['经营活动', '净额']
        )
        invest_col = self._find_column(
            cashflow,
            ['n_cashflow_inv_act'],
            ['投资活动', '净额']
        )

        if oper_col and invest_col:
            cashflow['fcff'] = cashflow[oper_col].fillna(0) + cashflow[invest_col].fillna(0)
        elif oper_col:
            capex_col = self._find_column(
                cashflow,
                ['c_paid_for_assets'],
                ['购建', '长期资产']
            )
            if capex_col:
                cashflow['fcff'] = cashflow[oper_col].fillna(0) - cashflow[capex_col].fillna(0).abs()
            else:
                cashflow['fcff'] = cashflow[oper_col].fillna(0) * 0.7
        else:
            return pd.DataFrame()

        # Tushare现金流数据单位为元，统一转换为万元
        cashflow['fcff'] = cashflow['fcff'] / 10000

        # 只保留年报数据(12月31日)并按年汇总
        if 'end_date' in cashflow.columns and 'fcff' in cashflow.columns:
            cashflow = cashflow[cashflow['end_date'].astype(str).str.contains('1231', na=False)]
            if cashflow.empty:
                return pd.DataFrame()

            cashflow['year'] = cashflow['end_date'].astype(str).str[:4].astype(int)
            fcf_data = cashflow.groupby('year').agg({
                'fcff': 'sum'
            }).reset_index()

            return fcf_data

        return pd.DataFrame()

    def _find_column(
        self,
        df: pd.DataFrame,
        english_names: List[str],
        chinese_keywords: List[str]
    ) -> Optional[str]:
        """在DataFrame中查找列名，优先英文名，回退中文关键字匹配"""
        for name in english_names:
            if name in df.columns:
                return name
        for col in df.columns:
            col_str = str(col)
            if all(kw in col_str for kw in chinese_keywords):
                return col
        return None

    def _calculate_ev_from_fcff(
        self,
        fcf_data: pd.DataFrame,
        wacc: float,
        terminal_growth_rate: float,
        forecast_years: int
    ) -> Tuple[float, float, list]:
        """方法1: FCFF法计算企业价值"""
        if fcf_data.empty:
            return 0, 0, []

        fcff_values = fcf_data['fcff'].values
        if len(fcff_values) >= 2:
            growth_rates = np.diff(fcff_values) / np.abs(fcff_values[:-1])
            growth_rates = growth_rates[np.isfinite(growth_rates)]
            growth_rates = growth_rates[(growth_rates > -0.5) & (growth_rates < 2.0)]
            avg_growth = np.mean(growth_rates) if len(growth_rates) > 0 else 0.05
        else:
            avg_growth = 0.05

        avg_growth = max(min(avg_growth, 0.20), -0.10)

        last_fcff = fcff_values[-1] if len(fcff_values) > 0 else 0
        forecast_cf_list = []

        for year in range(1, forecast_years + 1):
            forecast_cf_list.append(last_fcff * (1 + avg_growth) ** year)

        pv_cashflows = sum(
            cf / (1 + wacc) ** year
            for year, cf in enumerate(forecast_cf_list, 1)
        )

        if forecast_cf_list and wacc > terminal_growth_rate:
            terminal_fcff = forecast_cf_list[-1] * (1 + terminal_growth_rate)
            terminal_value = terminal_fcff / (wacc - terminal_growth_rate)
        else:
            terminal_value = 0

        pv_terminal = terminal_value / (1 + wacc) ** forecast_years
        ev = pv_cashflows + pv_terminal

        return ev, terminal_value, forecast_cf_list

    def _calculate_equity_from_fcfe(
        self,
        fcf_data: pd.DataFrame,
        ke: float,
        terminal_growth_rate: float,
        forecast_years: int,
        debt_ratio: float = 0.4
    ) -> Tuple[float, float, list]:
        """方法2: FCFE法计算股权价值

        FCFE = FCFF * (1 - debt_ratio) 简化估算，使用实际资产负债率
        """
        if fcf_data.empty:
            return 0, 0, []

        equity_ratio = 1 - debt_ratio
        fcfe_data = fcf_data.copy()
        fcfe_data['fcfe'] = fcfe_data['fcff'] * equity_ratio

        fcfe_values = fcfe_data['fcfe'].values
        if len(fcfe_values) >= 2:
            growth_rates = np.diff(fcfe_values) / np.abs(fcfe_values[:-1])
            growth_rates = growth_rates[np.isfinite(growth_rates)]
            growth_rates = growth_rates[(growth_rates > -0.5) & (growth_rates < 2.0)]
            avg_growth = np.mean(growth_rates) if len(growth_rates) > 0 else 0.05
        else:
            avg_growth = 0.05

        avg_growth = max(min(avg_growth, 0.20), -0.10)

        last_fcfe = fcfe_values[-1] if len(fcfe_values) > 0 else 0
        forecast_cf_list = []

        for year in range(1, forecast_years + 1):
            forecast_cf_list.append(last_fcfe * (1 + avg_growth) ** year)

        pv_cashflows = sum(
            cf / (1 + ke) ** year
            for year, cf in enumerate(forecast_cf_list, 1)
        )

        if forecast_cf_list and ke > terminal_growth_rate:
            terminal_fcfe = forecast_cf_list[-1] * (1 + terminal_growth_rate)
            terminal_value = terminal_fcfe / (ke - terminal_growth_rate)
        else:
            terminal_value = 0

        pv_terminal = terminal_value / (1 + ke) ** forecast_years
        equity_value = pv_cashflows + pv_terminal

        return equity_value, terminal_value, forecast_cf_list

    def _calculate_net_debt(self, financial_statements: Dict[str, pd.DataFrame]) -> float:
        """计算净债务

        使用Tushare标准英文字段名:
        - st_borr: 短期借款
        - lt_borr: 长期借款
        - non_cur_liab_due_1y: 一年内到期的非流动负债
        - money_cap: 货币资金
        """
        balance_sheet = financial_statements.get('balance_sheet', pd.DataFrame())

        if balance_sheet.empty:
            return 0

        if 'end_date' in balance_sheet.columns:
            bs_sorted = balance_sheet.sort_values('end_date', ascending=False)
            latest = bs_sorted.iloc[0]
        else:
            latest = balance_sheet.iloc[-1]

        st_borr = latest.get('st_borr', 0)
        if pd.isna(st_borr):
            st_borr = 0

        lt_borr = latest.get('lt_borr', 0)
        if pd.isna(lt_borr):
            lt_borr = 0

        non_cur_liab_due_1y = latest.get('non_cur_liab_due_1y', 0)
        if pd.isna(non_cur_liab_due_1y):
            non_cur_liab_due_1y = 0

        money_cap = latest.get('money_cap', 0)
        if pd.isna(money_cap):
            money_cap = 0

        # Tushare资产负债表数据单位为元，转换为万元
        debt = (float(st_borr) + float(lt_borr) + float(non_cur_liab_due_1y)) / 10000
        cash = float(money_cap) / 10000

        net_debt = debt - cash
        return max(net_debt, 0)

    def generate_valuation_summary(
        self,
        dcf_result: Dict[str, Any],
        shares_outstanding: float = None,
        current_price: float = None
    ) -> pd.DataFrame:
        """生成估值汇总表"""
        rows = []

        ev = dcf_result.get('method1_ev', 0)
        rows.append({
            '估值方法': 'FCFF法(企业价值)',
            '估值结果(万元)': round(ev, 2),
            '说明': 'EV = Σ FCFF/(1+WACC)^t + TV/(1+WACC)^n'
        })

        equity_fe = dcf_result.get('method2_equity_value', 0)
        rows.append({
            '估值结果(万元)': round(equity_fe, 2),
            '估值方法': 'FCFE法(股权价值)',
            '说明': '股权价值 = Σ FCFE/(1+Ke)^t + TV/(1+Ke)^n'
        })

        equity_indirect = dcf_result.get('method3_equity_value', 0)
        net_debt = dcf_result.get('net_debt', 0)
        rows.append({
            '估值结果(万元)': round(equity_indirect, 2),
            '估值方法': '间接法(股权价值)',
            '说明': f'股权价值 = EV - 净债务 ({net_debt:.2f}万元)'
        })

        df = pd.DataFrame(rows)

        # 每股价值(元) = 估值(万元) / 股本(万股)
        # 因为: 万元/万股 = 万元/(万股) = (万元*10000元/万元)/(万股*10000股/万股) = 元/股
        if shares_outstanding and shares_outstanding > 0:
            df['每股价值(元)'] = df['估值结果(万元)'] / shares_outstanding

            if current_price:
                df['估值空间'] = ((df['每股价值(元)'] / current_price) - 1) * 100
                df['估值空间'] = df['估值空间'].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "")

        return df
