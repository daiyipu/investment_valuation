#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第六章：公司估值
包含：
- 当前估值水平
- 历史FCF数据分析
- WACC/Ke计算过程
- DCF三种方法：FCFF法、FCFE法、间接法
- 历史估值对比
- 估值敏感性分析
- 公司估值评估
- 限售股解禁计划
"""

from typing import Dict, Any, List
from datetime import datetime
import pandas as pd
import numpy as np

from utils.dcf_calculator import DCFCalculator


class Chapter06Valuation:
    """公司估值分析章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any]):
        self.config = config
        self.data = data
        self.project_info = config.get('project', {})
        self.dcf_calculator = DCFCalculator()
        self._dcf_result = None

    def generate(self) -> List[Dict[str, Any]]:
        elements = []

        elements.append({'type': 'heading', 'content': '第六章 公司估值', 'level': 1})

        # 6.1 当前估值水平
        elements.append({'type': 'heading', 'content': '6.1 当前估值水平', 'level': 2})
        elements.extend(self._generate_current_valuation())

        # 6.2 历史自由现金流分析
        elements.append({'type': 'heading', 'content': '6.2 历史自由现金流分析', 'level': 2})
        elements.extend(self._generate_historical_fcf())

        # 6.3 DCF估值分析
        elements.append({'type': 'heading', 'content': '6.3 DCF估值分析', 'level': 2})
        elements.extend(self._generate_dcf_valuation())

        # 6.4 历史估值对比
        elements.append({'type': 'heading', 'content': '6.4 历史估值对比', 'level': 2})
        elements.extend(self._generate_historical_valuation())

        # 6.5 估值敏感性分析
        elements.append({'type': 'heading', 'content': '6.5 估值敏感性分析', 'level': 2})
        elements.extend(self._generate_sensitivity_analysis())

        # 6.6 估值综合评估
        elements.append({'type': 'heading', 'content': '6.6 估值综合评估', 'level': 2})
        elements.extend(self._generate_valuation_risk())

        # 6.7 限售股解禁计划
        elements.append({'type': 'heading', 'content': '6.7 限售股解禁计划', 'level': 2})
        elements.extend(self._generate_share_float())

        return elements

    # ==================== 6.1 当前估值水平 ====================

    def _generate_current_valuation(self) -> List[Dict[str, Any]]:
        elements = []
        vd = self.data.get('valuation_data', {})

        valuation_info = []
        if vd.get('current_price'):
            valuation_info.append({'指标': '最新收盘价', '数值': f"{vd['current_price']:.2f} 元"})
        if vd.get('pe_ratio') is not None:
            valuation_info.append({'指标': '市盈率(PE-TTM)', '数值': f"{vd['pe_ratio']:.2f}"})
        if vd.get('pb_ratio') is not None:
            valuation_info.append({'指标': '市净率(PB)', '数值': f"{vd['pb_ratio']:.2f}"})
        if vd.get('market_cap'):
            valuation_info.append({'指标': '总市值', '数值': f"{vd['market_cap']/10000:.2f} 亿元"})
        if vd.get('float_market_cap'):
            valuation_info.append({'指标': '流通市值', '数值': f"{vd['float_market_cap']/10000:.2f} 亿元"})
        if vd.get('turnover_rate') is not None:
            valuation_info.append({'指标': '换手率', '数值': f"{vd['turnover_rate']:.2f}%"})

        if not valuation_info:
            elements.append({'type': 'paragraph', 'content': '暂无估值数据'})
        else:
            elements.append({'type': 'table', 'headers': ['指标', '数值'], 'data': valuation_info})

        return elements

    # ==================== 6.2 历史FCF分析 ====================

    def _generate_historical_fcf(self) -> List[Dict[str, Any]]:
        """从现金流量表提取历史FCF数据并分析"""
        elements = []

        cashflow = self.data.get('financial_statements', {}).get('cashflow', pd.DataFrame())
        if cashflow.empty:
            elements.append({'type': 'paragraph', 'content': '暂无现金流数据，无法进行历史FCF分析'})
            return elements

        cashflow = cashflow.copy()

        oper_col = self._find_col(cashflow, ['n_cashflow_act'], ['经营活动', '净额'])
        invest_col = self._find_col(cashflow, ['n_cashflow_inv_act'], ['投资活动', '净额'])
        capex_col = self._find_col(cashflow, ['c_paid_for_assets'], ['购建', '长期资产'])

        if not oper_col:
            elements.append({'type': 'paragraph', 'content': '现金流数据缺少经营活动现金流字段'})
            return elements

        # 只保留年报并按end_date去重（同一期可能有多条记录）
        cashflow = cashflow[cashflow['end_date'].astype(str).str.contains('1231', na=False)]
        cashflow = cashflow.drop_duplicates(subset=['end_date'], keep='last')
        cashflow = cashflow.sort_values('end_date')

        rows = []
        fcf_values = []
        for _, row in cashflow.iterrows():
            year = str(row['end_date'])[:4]
            oper_cf = float(row[oper_col]) if pd.notna(row.get(oper_col)) else 0
            invest_cf = float(row[invest_col]) if invest_col and pd.notna(row.get(invest_col)) else 0
            capex = float(row[capex_col]) if capex_col and pd.notna(row.get(capex_col)) else 0

            if capex_col:
                fcf = oper_cf - abs(capex)
                fcf_method = '经营CF-资本支出'
            elif invest_col:
                fcf = oper_cf + invest_cf
                fcf_method = '经营CF+投资CF'
            else:
                fcf = oper_cf * 0.7
                fcf_method = '经营CF×0.7(估算)'

            fcf_values.append(fcf)
            rows.append({
                '年份': year,
                '经营活动CF(万元)': f"{oper_cf/10000:.2f}",
                '投资活动CF(万元)': f"{invest_cf/10000:.2f}",
                'FCF(万元)': f"{fcf/10000:.2f}",
                '计算方法': fcf_method if _ == len(cashflow) - 1 else '',
            })

        if rows:
            elements.append({
                'type': 'paragraph',
                'content': '自由现金流(FCF) = 经营活动现金流 - 资本支出，反映公司在满足再投资需要后剩余的现金流量。'
            })
            elements.append({
                'type': 'table',
                'headers': ['年份', '经营活动CF(万元)', '投资活动CF(万元)', 'FCF(万元)', '计算方法'],
                'data': rows
            })

            positive_fcfs = [v for v in fcf_values if v > 0]
            if len(positive_fcfs) > 0:
                avg_fcf = np.mean(positive_fcfs) / 10000
                trend_text = f'历史FCF正值年份 {len(positive_fcfs)}/{len(fcf_values)} 年，正值均值 {avg_fcf:.2f} 万元。'
            else:
                trend_text = '历史FCF均为负值，公司处于持续投入期。'

            if len(positive_fcfs) >= 2:
                first_pos = positive_fcfs[0]
                last_pos = positive_fcfs[-1]
                n_years = len(positive_fcfs) - 1
                if first_pos > 0:
                    cagr = (last_pos / first_pos) ** (1 / n_years) - 1
                    trend_text += f' FCF复合增长率(CAGR): {cagr*100:.2f}%。'

            elements.append({'type': 'paragraph', 'content': trend_text})

        return elements

    # ==================== 6.3 DCF估值分析 ====================

    def _generate_dcf_valuation(self) -> List[Dict[str, Any]]:
        elements = []

        cashflow = self.data.get('financial_statements', {}).get('cashflow', pd.DataFrame())
        if cashflow.empty:
            elements.append({'type': 'paragraph', 'content': '现金流数据不足，无法进行DCF估值'})
            return elements

        dcf_result = self._get_dcf_result()
        if not dcf_result:
            elements.append({'type': 'paragraph', 'content': 'DCF估值计算失败'})
            return elements

        # 6.3.1 估值方法说明
        elements.append({'type': 'heading', 'content': '6.3.1 DCF估值方法说明', 'level': 3})
        elements.append({
            'type': 'paragraph',
            'content': (
                'DCF（Discounted Cash Flow）估值法通过预测公司未来自由现金流，并以加权平均资本成本（WACC）折现至现值，得出公司的内在价值。\n\n'
                '本报告采用三种DCF方法进行交叉验证：\n'
                '1. FCFF法（公司自由现金流法）：以WACC折现FCFF得到企业价值(EV)，再减去净债务得到股权价值\n'
                '2. FCFE法（股权自由现金流法）：以股权成本Ke直接折现FCFE得到股权价值\n'
                '3. 间接法：EV - 净债务 = 股权价值'
            )
        })

        # 6.3.2 折现率计算
        elements.append({'type': 'heading', 'content': '6.3.2 折现率计算', 'level': 3})
        elements.extend(self._generate_discount_rate_section(dcf_result))

        # 6.3.3 FCF增长率估算
        elements.append({'type': 'heading', 'content': '6.3.3 FCF增长率估算', 'level': 3})
        elements.extend(self._generate_growth_section(dcf_result))

        # 6.3.4 方法一：FCFF法
        elements.append({'type': 'heading', 'content': '6.3.4 方法一：FCFF法（公司自由现金流法）', 'level': 3})
        elements.extend(self._generate_method_fcff(dcf_result))

        # 6.3.5 方法二：FCFE法
        elements.append({'type': 'heading', 'content': '6.3.5 方法二：FCFE法（股权自由现金流法）', 'level': 3})
        elements.extend(self._generate_method_fcfe(dcf_result))

        # 6.3.6 方法三：间接法
        elements.append({'type': 'heading', 'content': '6.3.6 方法三：间接法（EV-净债务）', 'level': 3})
        elements.extend(self._generate_method_indirect(dcf_result))

        # 6.3.7 三种方法汇总对比
        elements.append({'type': 'heading', 'content': '6.3.7 三种方法汇总对比', 'level': 3})
        elements.extend(self._generate_methods_summary(dcf_result))

        return elements

    def _generate_discount_rate_section(self, dcf_result) -> List[Dict[str, Any]]:
        """WACC和Ke计算过程"""
        elements = []

        wacc = dcf_result.get('wacc', 0)
        ke = dcf_result.get('ke', 0)
        debt_ratio = dcf_result.get('debt_ratio', 0.4)
        equity_ratio = 1 - debt_ratio
        kd = 0.05
        tax_rate = 0.25

        elements.append({
            'type': 'paragraph',
            'content': (
                f'WACC（加权平均资本成本）：用于FCFF法折现\n'
                f'WACC = 股权占比 × Ke + 债务占比 × Kd × (1-T)\n'
                f'     = {equity_ratio:.1%} × {ke*100:.2f}% + {debt_ratio:.1%} × {kd*100:.2f}% × (1-{tax_rate:.0%})\n'
                f'     = {wacc*100:.2f}%\n\n'
                f'Ke（股权成本）：用于FCFE法折现\n'
                f'Ke = Rf + β × (Rm - Rf)\n'
                f'   = {dcf_result.get("risk_free_rate", 0.03)*100:.2f}% + {dcf_result.get("beta", 1.0):.2f} × {dcf_result.get("market_risk_premium", 0.06)*100:.2f}%\n'
                f'   = {ke*100:.2f}%'
            )
        })

        params = [
            {'参数': '无风险利率(Rf)', '数值': f"{dcf_result.get('risk_free_rate', 0.03)*100:.2f}%"},
            {'参数': '市场风险溢价(Rm-Rf)', '数值': f"{dcf_result.get('market_risk_premium', 0.06)*100:.2f}%"},
            {'参数': 'Beta系数', '数值': f"{dcf_result.get('beta', 1.0):.2f}"},
            {'参数': '股权成本(Ke)', '数值': f"{ke*100:.2f}%"},
            {'参数': '债务成本(Kd)', '数值': f"{kd*100:.2f}%"},
            {'参数': '所得税率(T)', '数值': f"{tax_rate:.0%}"},
            {'参数': '资产负债率', '数值': f"{debt_ratio:.2%}"},
            {'参数': 'WACC', '数值': f"{wacc*100:.2f}%"},
        ]

        balance_sheet = self.data.get('financial_statements', {}).get('balance_sheet', pd.DataFrame())
        if not balance_sheet.empty:
            if 'end_date' in balance_sheet.columns:
                latest = balance_sheet.sort_values('end_date', ascending=False).iloc[0]
            else:
                latest = balance_sheet.iloc[-1]
            total_assets = latest.get('total_assets')
            total_liab = latest.get('total_liab')
            if total_assets is not None and pd.notna(total_assets):
                params.append({'参数': '总资产', '数值': f"{float(total_assets)/1e8:.2f} 亿元"})
            if total_liab is not None and pd.notna(total_liab):
                params.append({'参数': '总负债', '数值': f"{float(total_liab)/1e8:.2f} 亿元"})

        elements.append({'type': 'table', 'headers': ['参数', '数值'], 'data': params})
        return elements

    def _generate_growth_section(self, dcf_result) -> List[Dict[str, Any]]:
        elements = []

        fcf_data = self.dcf_calculator._prepare_fcf_data(
            self.data.get('financial_statements', {}),
            self.data.get('financial_indicators', pd.DataFrame())
        )

        if fcf_data.empty or len(fcf_data) < 2:
            elements.append({'type': 'paragraph', 'content': '历史FCF数据不足（少于2年），使用默认增长率5.0%'})
            return elements

        fcff_values = fcf_data['fcff'].values
        growth_rates = np.diff(fcff_values) / np.abs(fcff_values[:-1])
        growth_rates = growth_rates[np.isfinite(growth_rates)]
        growth_rates = growth_rates[(growth_rates > -0.5) & (growth_rates < 2.0)]
        if len(growth_rates) > 0:
            avg_growth = np.mean(growth_rates)
            growth_clamped = max(min(avg_growth, 0.20), -0.10)

            growth_rows = []
            for i, yr in enumerate(fcf_data['year'].values):
                growth_rows.append({'年份': str(yr), 'FCF(万元)': f"{fcff_values[i]:.2f}"})

            elements.append({'type': 'table', 'headers': ['年份', 'FCF(万元)'], 'data': growth_rows})

            elements.append({
                'type': 'paragraph',
                'content': (
                    f'历史FCF年均增长率: {avg_growth*100:.2f}%\n'
                    f'限制后采用增长率: {growth_clamped*100:.2f}%（限制在-10%~20%范围内）\n'
                    f'永续增长率: {dcf_result.get("terminal_growth_rate", 0.03)*100:.2f}%'
                )
            })

        return elements

    # ==================== 方法一：FCFF法 ====================

    def _generate_method_fcff(self, dcf_result) -> List[Dict[str, Any]]:
        elements = []

        wacc = dcf_result.get('wacc', 0)
        tg = dcf_result.get('terminal_growth_rate', 0.03)
        ev = dcf_result.get('method1_ev', 0)
        tv = dcf_result.get('method1_terminal_value', 0)
        cf_list = dcf_result.get('method1_cf_list', [])
        forecast_years = len(cf_list)

        if not cf_list:
            elements.append({'type': 'paragraph', 'content': 'FCFF法计算数据不足'})
            return elements

        elements.append({
            'type': 'paragraph',
            'content': (
                'FCFF法（Free Cash Flow to Firm）以公司整体自由现金流为基础，\n'
                '以WACC为折现率计算企业价值(EV)，再扣除净债务得到股权价值。\n'
                f'公式：EV = Σ FCFF/(1+WACC)^t + TV/(1+WACC)^n'
            )
        })

        # 逐年预测表
        elements.append({'type': 'paragraph', 'content': f'基于WACC={wacc*100:.2f}%，逐年预测FCFF并折现：'})

        pv_total = 0
        rows = []
        for i, fcf in enumerate(cf_list):
            year = i + 1
            discount_factor = 1 / (1 + wacc) ** year
            pv = fcf * discount_factor
            pv_total += pv
            rows.append({
                '年份': f'第{year}年',
                '预测FCFF(万元)': f"{fcf:.2f}",
                '折现因子': f"{discount_factor:.4f}",
                '现值(万元)': f"{pv:.2f}",
            })

        elements.append({'type': 'table', 'headers': ['年份', '预测FCFF(万元)', '折现因子', '现值(万元)'], 'data': rows})

        # 终值
        if wacc > tg:
            last_fcf = cf_list[-1]
            terminal_fcf = last_fcf * (1 + tg)
            pv_terminal = tv / (1 + wacc) ** forecast_years
            elements.append({
                'type': 'paragraph',
                'content': (
                    f'终值计算（Gordon Growth Model）：\n'
                    f'终值 = {last_fcf:.2f} × (1+{tg*100:.1f}%) / ({wacc*100:.2f}%-{tg*100:.1f}%) = {tv:.2f} 万元\n'
                    f'终值现值 = {tv:.2f} / (1+{wacc*100:.2f}%)^{forecast_years} = {pv_terminal:.2f} 万元'
                )
            })

        elements.append({
            'type': 'paragraph',
            'content': f'企业价值(EV) = 预测期现值 {pv_total:.2f} + 终值现值 {ev - pv_total:.2f} = {ev:.2f} 万元 = {ev/10000:.2f} 亿元'
        })

        return elements

    # ==================== 方法二：FCFE法 ====================

    def _generate_method_fcfe(self, dcf_result) -> List[Dict[str, Any]]:
        elements = []

        ke = dcf_result.get('ke', 0)
        tg = dcf_result.get('terminal_growth_rate', 0.03)
        equity_fe = dcf_result.get('method2_equity_value', 0)
        tv_fe = dcf_result.get('method2_terminal_value', 0)
        cf_list_fe = dcf_result.get('method2_cf_list', [])
        debt_ratio = dcf_result.get('debt_ratio', 0.4)

        if not cf_list_fe:
            elements.append({'type': 'paragraph', 'content': 'FCFE法计算数据不足'})
            return elements

        equity_ratio = 1 - debt_ratio
        elements.append({
            'type': 'paragraph',
            'content': (
                'FCFE法（Free Cash Flow to Equity）直接计算归属于股东的现金流，\n'
                f'以股权成本Ke={ke*100:.2f}%为折现率，直接得到股权价值。\n'
                f'FCFE = FCFF × 股权比例 = FCFF × {equity_ratio:.2%}'
            )
        })

        # 逐年预测表
        pv_total = 0
        rows = []
        for i, fcf in enumerate(cf_list_fe):
            year = i + 1
            discount_factor = 1 / (1 + ke) ** year
            pv = fcf * discount_factor
            pv_total += pv
            rows.append({
                '年份': f'第{year}年',
                '预测FCFE(万元)': f"{fcf:.2f}",
                '折现因子': f"{discount_factor:.4f}",
                '现值(万元)': f"{pv:.2f}",
            })

        elements.append({'type': 'table', 'headers': ['年份', '预测FCFE(万元)', '折现因子', '现值(万元)'], 'data': rows})

        # 终值
        if ke > tg:
            last_fcf = cf_list_fe[-1]
            pv_terminal = tv_fe / (1 + ke) ** len(cf_list_fe)
            elements.append({
                'type': 'paragraph',
                'content': (
                    f'终值 = {last_fcf:.2f} × (1+{tg*100:.1f}%) / ({ke*100:.2f}%-{tg*100:.1f}%) = {tv_fe:.2f} 万元\n'
                    f'终值现值 = {tv_fe:.2f} / (1+{ke*100:.2f}%)^{len(cf_list_fe)} = {pv_terminal:.2f} 万元'
                )
            })

        elements.append({
            'type': 'paragraph',
            'content': f'股权价值(FCFE法) = {equity_fe:.2f} 万元 = {equity_fe/10000:.2f} 亿元'
        })

        return elements

    # ==================== 方法三：间接法 ====================

    def _generate_method_indirect(self, dcf_result) -> List[Dict[str, Any]]:
        elements = []

        ev = dcf_result.get('method1_ev', 0)
        net_debt = dcf_result.get('net_debt', 0)
        equity_indirect = dcf_result.get('method3_equity_value', 0)

        elements.append({
            'type': 'paragraph',
            'content': (
                '间接法基于方法一的FCFF计算结果，从企业价值中扣除净债务得到股权价值。\n'
                f'公式：股权价值 = EV - 净债务'
            )
        })

        elements.append({
            'type': 'paragraph',
            'content': (
                f'企业价值(EV) = {ev:.2f} 万元（来自FCFF法）\n'
                f'净债务 = 有息负债 - 货币资金 = {net_debt:.2f} 万元 = {net_debt/10000:.2f} 亿元\n'
                f'股权价值 = {ev:.2f} - {net_debt:.2f} = {equity_indirect:.2f} 万元 = {equity_indirect/10000:.2f} 亿元'
            )
        })

        return elements

    # ==================== 三种方法汇总 ====================

    def _generate_methods_summary(self, dcf_result) -> List[Dict[str, Any]]:
        elements = []

        ev = dcf_result.get('method1_ev', 0)
        net_debt = dcf_result.get('net_debt', 0)
        equity_indirect = dcf_result.get('method3_equity_value', 0)
        equity_fe = dcf_result.get('method2_equity_value', 0)
        shares = self._get_shares_outstanding()
        current_price = self.data.get('valuation_data', {}).get('current_price')

        summary_rows = [
            {
                '估值方法': '方法一：FCFF法',
                '企业价值(万元)': f"{ev:,.2f}",
                '股权价值(万元)': '-',
                '折现率': f"WACC={dcf_result.get('wacc', 0)*100:.2f}%",
            },
            {
                '估值方法': '方法二：FCFE法',
                '企业价值(万元)': '-',
                '股权价值(万元)': f"{equity_fe:,.2f}",
                '折现率': f"Ke={dcf_result.get('ke', 0)*100:.2f}%",
            },
            {
                '估值方法': '方法三：间接法(EV-净债务)',
                '企业价值(万元)': f"{ev:,.2f}",
                '股权价值(万元)': f"{equity_indirect:,.2f}",
                '折现率': f"WACC={dcf_result.get('wacc', 0)*100:.2f}%",
            },
        ]

        elements.append({'type': 'table', 'headers': ['估值方法', '企业价值(万元)', '股权价值(万元)', '折现率'], 'data': summary_rows})

        # 每股价值
        if shares > 0:
            per_share_rows = [
                {'方法': '方法二：FCFE法', '每股价值(元)': f"{equity_fe/shares:.2f}"},
                {'方法': '方法三：间接法', '每股价值(元)': f"{equity_indirect/shares:.2f}"},
            ]

            if current_price:
                for r in per_share_rows:
                    val = float(r['每股价值(元)'])
                    margin = (val - current_price) / current_price * 100
                    r['当前股价(元)'] = f"{current_price:.2f}"
                    r['安全边际'] = f"{margin:+.1f}%"

            elements.append({'type': 'table', 'headers': list(per_share_rows[0].keys()), 'data': per_share_rows})

            # 加权平均
            avg_equity = (equity_fe + equity_indirect) / 2
            avg_per_share = avg_equity / shares
            text = f'综合两种方法的股权价值均值：{avg_equity/10000:.2f}亿元，每股{avg_per_share:.2f}元。'
            if current_price:
                avg_margin = (avg_per_share - current_price) / current_price * 100
                text += f'当前股价{current_price:.2f}元，安全边际{avg_margin:+.1f}%。'
            elements.append({'type': 'paragraph', 'content': text})

        return elements

    # ==================== 6.4 历史估值对比 ====================

    def _generate_historical_valuation(self) -> List[Dict[str, Any]]:
        elements = []
        hist_val = self.data.get('historical_valuation', pd.DataFrame())

        if hist_val.empty:
            elements.append({'type': 'paragraph', 'content': '暂无历史估值数据'})
            return elements

        hist_val = hist_val.copy()
        hist_val['year'] = hist_val['trade_date'].astype(str).str[:4]

        yearly_stats = []
        for year in sorted(hist_val['year'].unique()):
            yd = hist_val[hist_val['year'] == year]

            pe_col = 'pe_ttm' if 'pe_ttm' in yd.columns else 'pe'
            pe_vals = yd[pe_col].dropna() if pe_col in yd.columns else pd.Series()
            pb_vals = yd['pb'].dropna() if 'pb' in yd.columns else pd.Series()

            pe_vals = pe_vals[(pe_vals > 0) & (pe_vals < 500)]
            pb_vals = pb_vals[(pb_vals > 0) & (pb_vals < 50)]

            yearly_stats.append({
                '年份': year,
                'PE均值': f"{pe_vals.mean():.2f}" if len(pe_vals) > 0 else '-',
                'PE中位数': f"{pe_vals.median():.2f}" if len(pe_vals) > 0 else '-',
                'PE区间': f"{pe_vals.min():.1f}-{pe_vals.max():.1f}" if len(pe_vals) > 0 else '-',
                'PB均值': f"{pb_vals.mean():.2f}" if len(pb_vals) > 0 else '-',
                'PB中位数': f"{pb_vals.median():.2f}" if len(pb_vals) > 0 else '-',
            })

        if yearly_stats:
            elements.append({
                'type': 'table',
                'headers': ['年份', 'PE均值', 'PE中位数', 'PE区间', 'PB均值', 'PB中位数'],
                'data': yearly_stats
            })

        vd = self.data.get('valuation_data', {})
        current_pe = vd.get('pe_ratio')
        current_pb = vd.get('pb_ratio')

        percentile_texts = []
        if current_pe:
            all_pe = hist_val.get('pe_ttm', hist_val.get('pe', pd.Series())).dropna()
            all_pe = all_pe[(all_pe > 0) & (all_pe < 500)]
            if len(all_pe) > 0:
                pct = (all_pe < current_pe).sum() / len(all_pe) * 100
                percentile_texts.append(f'当前PE({current_pe:.2f})处于近3年历史的{pct:.1f}%分位')

        if current_pb:
            all_pb = hist_val['pb'].dropna() if 'pb' in hist_val.columns else pd.Series()
            all_pb = all_pb[(all_pb > 0) & (all_pb < 50)]
            if len(all_pb) > 0:
                pct = (all_pb < current_pb).sum() / len(all_pb) * 100
                percentile_texts.append(f'当前PB({current_pb:.2f})处于近3年历史的{pct:.1f}%分位')

        if percentile_texts:
            elements.append({'type': 'paragraph', 'content': '；'.join(percentile_texts) + '。'})

        return elements

    # ==================== 6.5 估值敏感性分析 ====================

    def _generate_sensitivity_analysis(self) -> List[Dict[str, Any]]:
        elements = []

        dcf_result = self._get_dcf_result()
        if not dcf_result:
            elements.append({'type': 'paragraph', 'content': '无DCF计算结果，无法进行敏感性分析'})
            return elements

        fcf_data = self.dcf_calculator._prepare_fcf_data(
            self.data.get('financial_statements', {}),
            self.data.get('financial_indicators', pd.DataFrame())
        )

        if fcf_data.empty:
            return elements

        fcff_values = fcf_data['fcff'].values
        last_fcff = fcff_values[-1] if len(fcff_values) > 0 else 0
        net_debt = dcf_result.get('net_debt', 0)
        shares = self._get_shares_outstanding()
        current_price = self.data.get('valuation_data', {}).get('current_price')

        if last_fcff <= 0 or shares <= 0:
            elements.append({'type': 'paragraph', 'content': 'FCF为负值或股本数据缺失，敏感性分析结果参考性有限'})
            return elements

        wacc_range = [0.07, 0.08, 0.09, 0.10, 0.11, 0.12]
        growth_range = [0.02, 0.05, 0.08, 0.10, 0.15, 0.20]
        tg = dcf_result.get('terminal_growth_rate', 0.03)
        forecast_years = 5

        matrix = []
        for g in growth_range:
            row = {}
            row['增长率'] = f'{g*100:.0f}%'
            for w in wacc_range:
                pv_cf = sum(
                    last_fcff * (1 + g) ** y / (1 + w) ** y
                    for y in range(1, forecast_years + 1)
                )
                last_proj = last_fcff * (1 + g) ** forecast_years
                if w > tg:
                    tv = last_proj * (1 + tg) / (w - tg)
                    pv_tv = tv / (1 + w) ** forecast_years
                else:
                    pv_tv = 0

                ev = pv_cf + pv_tv
                equity = ev - net_debt
                per_share = equity / shares if shares > 0 else 0

                key = f'WACC={w*100:.0f}%'
                if current_price:
                    margin = (per_share - current_price) / current_price * 100
                    row[key] = f'{per_share:.2f}元({margin:+.0f}%)'
                else:
                    row[key] = f'{per_share:.2f}元'
            matrix.append(row)

        wacc_headers = ['增长率'] + [f'WACC={w*100:.0f}%' for w in wacc_range]

        elements.append({
            'type': 'paragraph',
            'content': f'以下矩阵展示不同WACC和FCF增长率组合下的每股价值（当前股价: {current_price:.2f}元）' if current_price else '以下矩阵展示不同WACC和FCF增长率组合下的每股价值'
        })
        elements.append({'type': 'table', 'headers': wacc_headers, 'data': matrix})

        return elements

    # ==================== 6.6 估值综合评估 ====================

    def _generate_valuation_risk(self) -> List[Dict[str, Any]]:
        elements = []

        vd = self.data.get('valuation_data', {})
        pe_ratio = vd.get('pe_ratio')
        pb_ratio = vd.get('pb_ratio')

        dcf_result = self._get_dcf_result()

        risk_rows = []

        if pe_ratio:
            if pe_ratio < 0:
                pe_risk = '较高'
                pe_desc = f'PE为负值({pe_ratio:.1f})，公司处于亏损状态'
            elif pe_ratio < 20:
                pe_risk = '低'
                pe_desc = f'PE={pe_ratio:.1f}，估值处于较低水平'
            elif pe_ratio < 40:
                pe_risk = '中'
                pe_desc = f'PE={pe_ratio:.1f}，估值处于合理区间'
            else:
                pe_risk = '较高'
                pe_desc = f'PE={pe_ratio:.1f}，估值偏高需关注'
            risk_rows.append({'评估维度': 'PE估值', '评估结果': pe_desc, '风险等级': pe_risk})
        else:
            risk_rows.append({'评估维度': 'PE估值', '评估结果': '暂无数据', '风险等级': '待评估'})

        if pb_ratio:
            if pb_ratio < 1:
                pb_risk = '关注'
                pb_desc = f'PB={pb_ratio:.2f}<1，股价破净，需关注基本面'
            elif pb_ratio < 3:
                pb_risk = '中'
                pb_desc = f'PB={pb_ratio:.2f}，处于合理水平'
            else:
                pb_risk = '较高'
                pb_desc = f'PB={pb_ratio:.2f}，估值偏高'
            risk_rows.append({'评估维度': 'PB估值', '评估结果': pb_desc, '风险等级': pb_risk})

        if dcf_result:
            equity_fe = dcf_result.get('method2_equity_value', 0)
            equity_indirect = dcf_result.get('method3_equity_value', 0)
            shares = self._get_shares_outstanding()
            current_price = vd.get('current_price')

            if shares > 0:
                avg_equity = (equity_fe + equity_indirect) / 2
                intrinsic = avg_equity / shares
                dcf_desc = f'DCF每股价值={intrinsic:.2f}元（两种方法均值）'
                if current_price:
                    margin = (intrinsic - current_price) / current_price * 100
                    dcf_desc += f'，安全边际={margin:+.1f}%'
                    if margin > 15:
                        dcf_risk = '低'
                    elif margin > 0:
                        dcf_risk = '中'
                    else:
                        dcf_risk = '较高'
                else:
                    dcf_risk = '参考'
                risk_rows.append({'评估维度': 'DCF估值', '评估结果': dcf_desc, '风险等级': dcf_risk})

            market_cap = vd.get('market_cap')
            if market_cap and avg_equity != 0:
                deviation = (market_cap - avg_equity) / abs(avg_equity) * 100
                if deviation < -10:
                    dev_level = '低估'
                elif deviation > 10:
                    dev_level = '高估'
                else:
                    dev_level = '合理'
                risk_rows.append({
                    '评估维度': '市值/DCF偏离度',
                    '评估结果': f'市值{market_cap/10000:.2f}亿 vs DCF {avg_equity/10000:.2f}亿 ({deviation:+.1f}%)',
                    '风险等级': dev_level
                })

        elements.append({'type': 'table', 'headers': ['评估维度', '评估结果', '风险等级'], 'data': risk_rows})
        return elements

    # ==================== 6.7 限售股解禁计划 ====================

    def _generate_share_float(self) -> List[Dict[str, Any]]:
        elements = []
        share_float = self.data.get('share_float', pd.DataFrame())

        if share_float.empty:
            elements.append({'type': 'paragraph', 'content': '暂无限售股解禁数据。'})
            return elements

        share_float = share_float.copy()
        if 'float_date' in share_float.columns:
            share_float = share_float.sort_values('float_date')

        today = datetime.now().strftime('%Y%m%d')
        if 'float_date' in share_float.columns:
            future = share_float[share_float['float_date'].astype(str) >= today]
        else:
            future = share_float

        table_data = []
        for _, row in future.head(10).iterrows():
            float_date = str(row.get('float_date', ''))
            if len(float_date) >= 8:
                float_date = f"{float_date[:4]}-{float_date[4:6]}-{float_date[6:8]}"

            float_share = row.get('float_share', None)
            float_ratio = row.get('float_ratio', None)
            holder_type = row.get('holder_type', '')
            type_map = {'0': '其他', '1': '高管', '2': '个人', '3': '机构'}

            table_data.append({
                '解禁日期': float_date,
                '解禁数量(股)': f"{float(float_share):,.0f}" if pd.notna(float_share) else '-',
                '占总股本比例': f"{float(float_ratio):.2f}%" if pd.notna(float_ratio) else '-',
                '股东类型': type_map.get(str(holder_type), str(holder_type)),
            })

        if table_data:
            elements.append({'type': 'paragraph', 'content': '以下为未来限售股解禁计划：'})
            elements.append({
                'type': 'table',
                'headers': ['解禁日期', '解禁数量(股)', '占总股本比例', '股东类型'],
                'data': table_data
            })

            total_float_ratio = sum(
                float(r['占总股本比例'].rstrip('%'))
                for r in table_data if r['占总股本比例'] != '-'
            )
            if total_float_ratio > 20:
                text = f"未来解禁股份合计占总股本{total_float_ratio:.2f}%，解禁压力较大，需关注股价冲击风险。"
            elif total_float_ratio > 5:
                text = f"未来解禁股份合计占总股本{total_float_ratio:.2f}%，解禁压力适中。"
            else:
                text = f"未来解禁股份合计占总股本{total_float_ratio:.2f}%，解禁压力较小。"
            elements.append({'type': 'paragraph', 'content': text})
        else:
            elements.append({'type': 'paragraph', 'content': '当前暂无未来限售股解禁计划。'})

        return elements

    # ==================== 辅助方法 ====================

    def _get_dcf_result(self) -> Dict[str, Any]:
        if self._dcf_result is None:
            financial_statements = self.data.get('financial_statements', {})
            financial_indicators = self.data.get('financial_indicators', pd.DataFrame())
            cashflow = financial_statements.get('cashflow', pd.DataFrame())

            if not cashflow.empty:
                self._dcf_result = self.dcf_calculator.calculate_dcf_valuation(
                    financial_statements=financial_statements,
                    financial_indicators=financial_indicators
                )
        return self._dcf_result

    def _get_shares_outstanding(self) -> float:
        """获取总股本(万股)"""
        vd = self.data.get('valuation_data', {})
        market_cap = vd.get('market_cap')
        current_price = vd.get('current_price')

        if market_cap and current_price and current_price > 0:
            return float(market_cap) / float(current_price)

        balance_sheet = self.data.get('financial_statements', {}).get('balance_sheet', pd.DataFrame())
        if not balance_sheet.empty:
            if 'end_date' in balance_sheet.columns:
                latest = balance_sheet.sort_values('end_date', ascending=False).iloc[0]
            else:
                latest = balance_sheet.iloc[-1]

            for field in ['total_share', 'capital_stk']:
                value = latest.get(field)
                if value is not None and pd.notna(value) and value > 0:
                    return float(value) / 10000

        return 100000

    def _find_col(self, df, english_names, chinese_keywords):
        for name in english_names:
            if name in df.columns:
                return name
        for col in df.columns:
            if all(kw in str(col) for kw in chinese_keywords):
                return col
        return None
