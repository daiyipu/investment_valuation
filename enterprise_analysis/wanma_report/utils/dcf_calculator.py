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
from typing import Dict, Any, Optional, Tuple


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

        # 计算折现率
        wacc = self._calculate_wacc(risk_free_rate, market_risk_premium, beta, financial_statements)
        ke = self._calculate_ke(risk_free_rate, market_risk_premium, beta)

        result['wacc'] = wacc
        result['ke'] = ke
        result['risk_free_rate'] = risk_free_rate
        result['market_risk_premium'] = market_risk_premium
        result['beta'] = beta
        result['terminal_growth_rate'] = terminal_growth_rate

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

            # 方法2: FCFE法计算股权价值
            equity_value_fe, terminal_value_fe, cf_list_fe = self._calculate_equity_from_fcfe(
                fcf_data, ke, terminal_growth_rate, forecast_years
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

    def _calculate_wacc(
        self,
        risk_free_rate: float,
        market_risk_premium: float,
        beta: float,
        financial_statements: Dict[str, pd.DataFrame]
    ) -> float:
        """计算WACC

        Args:
            risk_free_rate: 无风险利率
            market_risk_premium: 市场风险溢价
            beta: Beta系数
            financial_statements: 财务报表

        Returns:
            WACC
        """
        # 权益成本 Ke = Rf + β * (Rm - Rf)
        ke = risk_free_rate + beta * market_risk_premium

        # 债务成本 (假设为5%)
        kd = 0.05

        # 计算债务和权益比例
        balance_sheet = financial_statements.get('balance_sheet', pd.DataFrame())
        if balance_sheet.empty:
            # 默认假设60%权益，40%债务
            equity_ratio = 0.6
            debt_ratio = 0.4
        else:
            # 从资产负债表中获取数据
            # 这里简化处理
            equity_ratio = 0.6
            debt_ratio = 0.4

        # 税率 (假设25%)
        tax_rate = 0.25

        # WACC = Ke * E/(E+D) + Kd * (1-T) * D/(E+D)
        wacc = ke * equity_ratio + kd * (1 - tax_rate) * debt_ratio

        return wacc

    def _calculate_ke(
        self,
        risk_free_rate: float,
        market_risk_premium: float,
        beta: float
    ) -> float:
        """计算权益成本Ke

        Args:
            risk_free_rate: 无风险利率
            market_risk_premium: 市场风险溢价
            beta: Beta系数

        Returns:
            Ke
        """
        return risk_free_rate + beta * market_risk_premium

    def _prepare_fcf_data(
        self,
        financial_statements: Dict[str, pd.DataFrame],
        financial_indicators: pd.DataFrame
    ) -> pd.DataFrame:
        """准备FCF数据

        Args:
            financial_statements: 财务报表
            financial_indicators: 财务指标

        Returns:
            FCF历史数据
        """
        cashflow = financial_statements.get('cashflow', pd.DataFrame())
        income = financial_statements.get('income_statement', pd.DataFrame())
        balance_sheet = financial_statements.get('balance_sheet', pd.DataFrame())

        if cashflow.empty:
            return pd.DataFrame()

        # 提取FCFF和FCFE数据
        # FCFF = 经营活动现金流 + 税后利息支出 - 资本支出
        # FCFE = 经营活动现金流 - 资本支出 + 净借款

        cashflow = cashflow.copy()

        # 简化处理：直接使用现金流量表中的自由现金流
        if 'c、资金活动产生的现金流量净额' in cashflow.columns:
            cashflow['fcff'] = cashflow['c、资金活动产生的现金流量净额'].fillna(0)

        # 如果有经营活动现金流
        oper_cf_cols = [c for c in cashflow.columns if '经营活动' in str(c) and '净额' in str(c)]
        invest_cf_cols = [c for c in cashflow.columns if '投资活动' in str(c) and '净额' in str(c)]

        if oper_cf_cols and invest_cf_cols:
            cashflow['fcff'] = cashflow[oper_cf_cols[0]].fillna(0) + cashflow[invest_cf_cols[0]].fillna(0)

        # 提取年份
        if 'end_date' in cashflow.columns and 'fcff' in cashflow.columns:
            cashflow['year'] = cashflow['end_date'].astype(str).str[:4].astype(int)

            # 按年份汇总
            fcf_data = cashflow.groupby('year').agg({
                'fcff': 'sum'
            }).reset_index()

            return fcf_data

        return pd.DataFrame()

    def _calculate_ev_from_fcff(
        self,
        fcf_data: pd.DataFrame,
        wacc: float,
        terminal_growth_rate: float,
        forecast_years: int
    ) -> Tuple[float, float, list]:
        """方法1: FCFF法计算企业价值

        Args:
            fcf_data: 历史FCF数据
            wacc: 加权平均资本成本
            terminal_growth_rate: 永续增长率
            forecast_years: 预测年数

        Returns:
            (企业价值, 终值, 现金流列表)
        """
        if fcf_data.empty:
            return 0, 0, []

        # 计算历史平均增长率
        fcff_values = fcf_data['fcff'].values
        if len(fcff_values) >= 2:
            growth_rates = np.diff(fcff_values) / np.abs(fcff_values[:-1])
            avg_growth = np.mean(growth_rates)
        else:
            avg_growth = 0.05

        # 限制增长率在合理范围
        avg_growth = max(min(avg_growth, 0.20), -0.10)

        # 预测未来FCFF
        last_fcff = fcff_values[-1] if len(fcff_values) > 0 else 0
        forecast_cf_list = []

        for year in range(1, forecast_years + 1):
            forecast_fcff = last_fcff * (1 + avg_growth) ** year
            forecast_cf_list.append(forecast_fcff)

        # 计算预测期现金流现值
        pv_cashflows = 0
        for year, cf in enumerate(forecast_cf_list, 1):
            pv = cf / (1 + wacc) ** year
            pv_cashflows += pv

        # 计算终值
        if len(forecast_cf_list) > 0:
            terminal_fcff = forecast_cf_list[-1] * (1 + terminal_growth_rate)
            terminal_value = terminal_fcff / (wacc - terminal_growth_rate)
        else:
            terminal_value = 0

        # 终值现值
        pv_terminal = terminal_value / (1 + wacc) ** forecast_years

        # 企业价值
        ev = pv_cashflows + pv_terminal

        return ev, terminal_value, forecast_cf_list

    def _calculate_equity_from_fcfe(
        self,
        fcf_data: pd.DataFrame,
        ke: float,
        terminal_growth_rate: float,
        forecast_years: int
    ) -> Tuple[float, float, list]:
        """方法2: FCFE法计算股权价值

        Args:
            fcf_data: 历史FCF数据
            ke: 权益资本成本
            terminal_growth_rate: 永续增长率
            forecast_years: 预测年数

        Returns:
            (股权价值, 终值, 现金流列表)
        """
        if fcf_data.empty:
            return 0, 0, []

        # FCFE ≈ FCFF - 税后利息支出 + 净借款
        # 简化处理: FCFE = FCFF * (1 - 债务比例)
        debt_ratio = 0.4

        fcfe_data = fcf_data.copy()
        fcfe_data['fcfe'] = fcfe_data['fcff'] * (1 - debt_ratio)

        # 计算历史平均增长率
        fcfe_values = fcfe_data['fcfe'].values
        if len(fcfe_values) >= 2:
            growth_rates = np.diff(fcfe_values) / np.abs(fcfe_values[:-1])
            avg_growth = np.mean(growth_rates)
        else:
            avg_growth = 0.05

        avg_growth = max(min(avg_growth, 0.20), -0.10)

        # 预测未来FCFE
        last_fcfe = fcfe_values[-1] if len(fcfe_values) > 0 else 0
        forecast_cf_list = []

        for year in range(1, forecast_years + 1):
            forecast_fcfe = last_fcfe * (1 + avg_growth) ** year
            forecast_cf_list.append(forecast_fcfe)

        # 计算预测期现金流现值
        pv_cashflows = 0
        for year, cf in enumerate(forecast_cf_list, 1):
            pv = cf / (1 + ke) ** year
            pv_cashflows += pv

        # 计算终值
        if len(forecast_cf_list) > 0:
            terminal_fcfe = forecast_cf_list[-1] * (1 + terminal_growth_rate)
            terminal_value = terminal_fcfe / (ke - terminal_growth_rate)
        else:
            terminal_value = 0

        # 终值现值
        pv_terminal = terminal_value / (1 + ke) ** forecast_years

        # 股权价值
        equity_value = pv_cashflows + pv_terminal

        return equity_value, terminal_value, forecast_cf_list

    def _calculate_net_debt(self, financial_statements: Dict[str, pd.DataFrame]) -> float:
        """计算净债务

        Args:
            financial_statements: 财务报表

        Returns:
            净债务
        """
        balance_sheet = financial_statements.get('balance_sheet', pd.DataFrame())

        if balance_sheet.empty:
            return 0

        # 净债务 = 债务 - 现金
        # 简化处理：假设债务为短期借款 + 长期借款，现金为货币资金

        debt = 0
        cash = 0

        # 查找相关科目 (使用中文列名)
        for col in balance_sheet.columns:
            col_str = str(col)
            if '借款' in col_str and '一年' not in col_str and '应付' not in col_str:
                if '短期借款' in col_str:
                    debt += balance_sheet[col].fillna(0).iloc[-1] if len(balance_sheet) > 0 else 0
            elif '长期借款' in col_str:
                debt += balance_sheet[col].fillna(0).iloc[-1] if len(balance_sheet) > 0 else 0
            elif '货币资金' in col_str:
                cash += balance_sheet[col].fillna(0).iloc[-1] if len(balance_sheet) > 0 else 0

        # 如果没有找到具体科目，返回一个估算值
        if debt == 0:
            # 使用财务指标估算
            return 0

        net_debt = debt - cash
        return max(net_debt, 0)

    def generate_valuation_summary(
        self,
        dcf_result: Dict[str, Any],
        shares_outstanding: float = None,
        current_price: float = None
    ) -> pd.DataFrame:
        """生成估值汇总表

        Args:
            dcf_result: DCF估值结果
            shares_outstanding: 总股本(万股)
            current_price: 当前股价

        Returns:
            估值汇总DataFrame
        """
        rows = []

        # 方法1: FCFF法
        ev = dcf_result.get('method1_ev', 0)
        rows.append({
            '估值方法': 'FCFF法(企业价值)',
            '估值结果(万元)': round(ev, 2),
            '说明': 'EV = Σ FCFF/(1+WACC)^t + TV/(1+WACC)^n'
        })

        # 方法2: FCFE法
        equity_fe = dcf_result.get('method2_equity_value', 0)
        rows.append({
            '估值结果(万元)': round(equity_fe, 2),
            '估值方法': 'FCFE法(股权价值)',
            '说明': '股权价值 = Σ FCFE/(1+Ke)^t + TV/(1+Ke)^n'
        })

        # 方法3: 间接法
        equity_indirect = dcf_result.get('method3_equity_value', 0)
        net_debt = dcf_result.get('net_debt', 0)
        rows.append({
            '估值结果(万元)': round(equity_indirect, 2),
            '估值方法': '间接法(股权价值)',
            '说明': f'股权价值 = EV - 净债务 ({net_debt:.2f}万元)'
        })

        df = pd.DataFrame(rows)

        # 如果有股本数据，计算每股价值
        if shares_outstanding and shares_outstanding > 0:
            df['每股价值(元)'] = df['估值结果(万元)'] / shares_outstanding * 10000

            if current_price:
                df['估值空间'] = ((df['每股价值(元)'] / current_price) - 1) * 100
                df['估值空间'] = df['估值空间'].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "")

        return df
