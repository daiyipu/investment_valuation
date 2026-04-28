#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第四章：DCF绝对估值
WACC计算明细、历史FCF分析、三方法DCF投影、每股价值汇总、敏感性矩阵热力图
"""

import os
import sys
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

import module_utils as utils

# Import DCFCalculator from utils
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PROJECT_DIR = os.path.dirname(os.path.dirname(_SCRIPT_DIR))
if _PROJECT_DIR not in sys.path:
    sys.path.insert(0, _PROJECT_DIR)

from utils.dcf_calculator import DCFCalculator


def _safe_val(row, field):
    """Safely extract a numeric value from a DataFrame row."""
    v = row.get(field)
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    return float(v)


def _get_annual_periods(df, n=5):
    """Get the most recent n annual reporting periods (ending 1231)."""
    if df is None or df.empty or 'end_date' not in df.columns:
        return []
    df_copy = df.copy()
    df_copy['end_date_str'] = df_copy['end_date'].astype(str)
    annual = df_copy[df_copy['end_date_str'].str.contains('1231', na=False)]
    periods = sorted(annual['end_date_str'].unique(), reverse=True)
    return list(periods[:n])


def _get_row_for_period(df, period):
    """Get the data row for a specific period."""
    if df is None or df.empty:
        return None
    df_copy = df.copy()
    df_copy['end_date_str'] = df_copy['end_date'].astype(str)
    period_data = df_copy[df_copy['end_date_str'] == period]
    if period_data.empty:
        return None
    return period_data.iloc[0]


def generate_chapter(context):
    """第四章：DCF绝对估值"""
    document = context['document']
    stock_code = context['stock_code']
    stock_name = context['stock_name']
    font_prop = context['font_prop']
    financial_statements = context.get('financial_statements', {})
    wacc_result = context.get('wacc_result', {})
    daily_basic = context.get('daily_basic', {})
    IMAGES_DIR = context['IMAGES_DIR']

    cashflow_df = financial_statements.get('cashflow', pd.DataFrame())
    bs_df = financial_statements.get('balancesheet', pd.DataFrame())
    income_df = financial_statements.get('income', pd.DataFrame())

    current_price = daily_basic.get('close', 0)
    if isinstance(current_price, str) or current_price is None:
        current_price = 0
    total_share = daily_basic.get('total_share', 0)
    if isinstance(total_share, str) or total_share is None:
        total_share = 0
    # Fallback: try to get total_shares from balancesheet if daily_basic missing
    if not total_share:
        if not bs_df.empty and 'total_hldr_eqy_inc_min_int' in bs_df.columns:
            # Use total equity as approximation (yuan → wan shares needs price)
            pass
        # Try financial_indicators
        fi = context.get('financial_indicators', pd.DataFrame())
        if not fi.empty and 'total_share' in fi.columns:
            ts_val = fi['total_share'].dropna()
            if not ts_val.empty:
                total_share = float(ts_val.iloc[0])
    if not total_share:
        print(f"  警告: 总股本数据缺失，每股价值计算将不可用")

    utils.add_title(document, '第四章 DCF绝对估值', level=1)

    # ==================== 4.1 WACC计算明细表 ====================
    utils.add_title(document, '4.1 WACC计算明细', level=2)

    if wacc_result:
        wacc_val = wacc_result.get('wacc', 0)
        ke = wacc_result.get('ke', wacc_result.get('cost_of_equity', 0))
        kd_pretax = wacc_result.get('kd_pretax', 0)
        kd_aftertax = wacc_result.get('kd_aftertax', wacc_result.get('cost_of_debt', 0))
        beta = wacc_result.get('adopted_beta', wacc_result.get('beta', 0))
        stock_beta = wacc_result.get('stock_beta', beta)
        industry_beta = wacc_result.get('industry_beta', beta)
        risk_free_rate = wacc_result.get('risk_free_rate', 0.0185)
        market_risk_premium = wacc_result.get('market_risk_premium', 0.0615)
        tax_rate = wacc_result.get('tax_rate', 0.25)
        equity_ratio = wacc_result.get('equity_ratio', 0)
        debt_ratio = wacc_result.get('debt_ratio', 0)
        market_cap = wacc_result.get('market_cap', 0)
        total_debt = wacc_result.get('total_debt', 0)

        wacc_headers = ['参数', '数值', '说明']
        wacc_data = [
            ['无风险利率(Rf)', f'{risk_free_rate:.4f}', '10年期国债收益率'],
            ['市场风险溢价(Rm-Rf)', f'{market_risk_premium:.4f}', '沪深300预期收益率-Rf'],
            ['个股Beta', f'{stock_beta:.4f}', '基于日收益率回归计算'],
            ['行业Beta', f'{industry_beta:.4f}', 'SW三级同行Beta中位数'],
            ['采用Beta', f'{beta:.4f}', '采用行业Beta'],
            ['股权成本(Ke)', f'{ke:.4f}', f'Rf + Beta × (Rm-Rf)'],
            ['税前债务成本(Kd)', f'{kd_pretax:.4f}', f'Rf × (1+债务溢价)'],
            ['税率', f'{tax_rate:.4f}', '企业所得税率'],
            ['税后债务成本', f'{kd_aftertax:.4f}', 'Kd × (1-税率)'],
            ['股权占比', f'{equity_ratio:.2%}', '市值/(市值+有息负债)'],
            ['债务占比', f'{debt_ratio:.2%}', '有息负债/(市值+有息负债)'],
            ['WACC', f'{wacc_val:.4f}', 'Ke×股权占比 + Kd×(1-t)×债务占比'],
        ]
        utils.add_table_data(document, wacc_headers, wacc_data)

        utils.add_paragraph(
            document,
            f'公司WACC为{wacc_val * 100:.2f}%，其中股权成本{ke * 100:.2f}%，'
            f'税后债务成本{kd_aftertax * 100:.2f}%，股权占比{equity_ratio:.1%}，'
            f'债务占比{debt_ratio:.1%}。'
        )
    else:
        utils.add_paragraph(document, '暂无WACC计算结果。')

    # ==================== 4.2 历史FCF分析 ====================
    utils.add_title(document, '4.2 历史自由现金流(FCFF)分析', level=2)
    utils.add_paragraph(document,
        'FCFF（公司自由现金流）= 经营活动现金流 + 利息支出×(1-T) - 资本支出。'
        '在中国会计准则下，利息支出归入经营活动现金流出，因此需加回税后利息以得到归属全部资本提供者的自由现金流。')

    income_df = financial_statements.get('income', pd.DataFrame())
    # Build interest lookup: year -> interest expense
    interest_by_year = {}
    if income_df is not None and not income_df.empty and 'end_date' in income_df.columns:
        inc_annual = income_df[income_df['end_date'].astype(str).str.contains('1231', na=False)].copy()
        inc_annual = inc_annual.drop_duplicates(subset='end_date', keep='last')
        for _, inc_row in inc_annual.iterrows():
            yr = str(inc_row['end_date'])[:4]
            int_exp = inc_row.get('fin_exp_int_exp', 0)
            if int_exp is None or (isinstance(int_exp, float) and np.isnan(int_exp)):
                int_exp = 0
            interest_by_year[yr] = float(int_exp)

    tax_rate = wacc_result.get('parameters', {}).get('tax_rate', 0.25) if wacc_result else 0.25

    fcf_records = []
    if not cashflow_df.empty:
        periods = _get_annual_periods(cashflow_df, 5)
        if periods:
            fcf_headers = ['年度', '经营现金流(万元)', '利息支出(万元)', '税后利息(万元)', '资本支出(万元)', 'FCFF(万元)']
            fcf_data = []

            for p in periods:
                r = _get_row_for_period(cashflow_df, p)
                if r is None:
                    continue

                year = str(p)[:4]
                ocf = _safe_val(r, 'n_cashflow_act') or 0
                capex = _safe_val(r, 'c_pay_acq_const_fiolta')
                inv_cf = _safe_val(r, 'n_cashflow_inv_act') or 0

                # After-tax interest add-back
                int_exp = interest_by_year.get(year, 0)
                after_tax_int = int_exp * (1 - tax_rate)

                if capex is not None and capex != 0:
                    fcf = ocf + after_tax_int - capex
                elif inv_cf != 0:
                    capex = abs(inv_cf)
                    fcf = ocf + after_tax_int - capex
                else:
                    capex = 0
                    fcf = (ocf + after_tax_int) * 0.7

                fcf_records.append({
                    'year': year,
                    'ocf': ocf / 10000,
                    'capex': capex / 10000 if capex else 0,
                    'fcf': fcf / 10000,
                    'interest_exp': int_exp / 10000,
                    'after_tax_int': after_tax_int / 10000,
                })

                fcf_data.append([
                    year,
                    f'{ocf / 10000:,.2f}',
                    f'{int_exp / 10000:,.2f}' if int_exp else '-',
                    f'{after_tax_int / 10000:,.2f}' if after_tax_int else '-',
                    f'{capex / 10000:,.2f}' if capex else '-',
                    f'{fcf / 10000:,.2f}',
                ])

            if fcf_data:
                utils.add_table_data(document, fcf_headers, fcf_data)

                # Summary
                fcf_values = [rec['fcf'] for rec in fcf_records]
                if fcf_values:
                    avg_fcf = np.mean(fcf_values)
                    utils.add_paragraph(
                        document,
                        f'近{len(fcf_values)}年平均FCF为{avg_fcf:,.2f}万元。'
                    )
            else:
                utils.add_paragraph(document, '暂无年度FCF数据。')
        else:
            utils.add_paragraph(document, '暂无年度现金流量表数据。')
    else:
        utils.add_paragraph(document, '暂无现金流量表数据。')

    # ==================== 4.3 三方法DCF投影 ====================
    utils.add_title(document, '4.3 三方法DCF估值', level=2)

    # Prepare DCF parameters from wacc_result
    dcf_params = {
        'risk_free_rate': wacc_result.get('risk_free_rate', 0.0185),
        'market_risk_premium': wacc_result.get('market_risk_premium', 0.0615),
        'beta': wacc_result.get('adopted_beta', wacc_result.get('beta', 1.0)),
        'kd': wacc_result.get('kd_pretax', wacc_result.get('risk_free_rate', 0.0185) * 1.5),
        'tax_rate': wacc_result.get('tax_rate', 0.25),
        'terminal_growth_rate': 0.025,
        'forecast_years': 10,
    }

    calculator = DCFCalculator()
    dcf_result = calculator.calculate_dcf_valuation(
        financial_statements=financial_statements,
        financial_indicators=context.get('financial_indicators', pd.DataFrame()),
        params=dcf_params,
    )

    if 'error' in dcf_result:
        utils.add_paragraph(document, f'DCF计算失败：{dcf_result["error"]}')
    else:
        # Display DCF parameters used
        utils.add_paragraph(document, 'DCF模型参数：', bold=True)

        param_headers = ['参数', '数值']
        param_data = [
            ['预测期(年)', str(dcf_result.get('forecast_years', 10))],
            ['WACC', f'{dcf_result.get("wacc", 0) * 100:.2f}%'],
            ['股权成本(Ke)', f'{dcf_result.get("ke", 0) * 100:.2f}%'],
            ['FCF增长率', f'{dcf_result.get("growth_rate", 0) * 100:.2f}%'],
            ['永续增长率', f'{dcf_result.get("terminal_growth_rate", 0) * 100:.2f}%'],
            ['股权占比', f'{dcf_result.get("equity_ratio", 0):.2%}'],
            ['债务占比', f'{dcf_result.get("debt_ratio", 0):.2%}'],
            ['净债务(万元)', f'{dcf_result.get("net_debt", 0):,.2f}'],
        ]
        utils.add_table_data(document, param_headers, param_data)

        utils.add_paragraph(document, '')

        # ---- Method 1: FCFF ----
        utils.add_paragraph(document, '方法一：FCFF（公司自由现金流法）', bold=True)

        ev_1 = dcf_result.get('method1_ev', 0)
        pv_forecasts_1 = dcf_result.get('method1_pv_forecasts', 0)
        tv_1 = dcf_result.get('method1_terminal_value', 0)
        pv_tv_1 = dcf_result.get('method1_pv_terminal', 0)

        fcff_headers = ['项目', '金额(万元)']
        fcff_data = [
            ['预测期FCF现值', f'{pv_forecasts_1:,.2f}'],
            ['终值', f'{tv_1:,.2f}'],
            ['终值现值', f'{pv_tv_1:,.2f}'],
            ['企业价值(EV)', f'{ev_1:,.2f}'],
            ['减：净债务', f'{dcf_result.get("net_debt", 0):,.2f}'],
            ['股权价值', f'{ev_1 - dcf_result.get("net_debt", 0):,.2f}'],
        ]
        utils.add_table_data(document, fcff_headers, fcff_data)

        # Show projected FCF details
        cf_list_1 = dcf_result.get('method1_cf_list', [])
        if cf_list_1:
            proj_headers = ['年份', 'FCF(万元)', '折现因子', '现值(万元)']
            proj_data = []
            for cf in cf_list_1:
                proj_data.append([
                    f'第{cf["year"]}年',
                    f'{cf["fcf"]:,.2f}',
                    f'{cf["discount_factor"]:.4f}',
                    f'{cf["pv"]:,.2f}',
                ])
            utils.add_table_data(document, proj_headers, proj_data)

        # ---- Method 2: FCFE ----
        utils.add_paragraph(document, '')
        utils.add_paragraph(document, '方法二：FCFE（股权自由现金流法）', bold=True)

        equity_2 = dcf_result.get('method2_equity_value', 0)
        pv_forecasts_2 = dcf_result.get('method2_pv_forecasts', 0)
        tv_2 = dcf_result.get('method2_terminal_value', 0)
        pv_tv_2 = dcf_result.get('method2_pv_terminal', 0)

        fcfe_headers = ['项目', '金额(万元)']
        fcfe_data = [
            ['预测期FCFE现值', f'{pv_forecasts_2:,.2f}'],
            ['终值', f'{tv_2:,.2f}'],
            ['终值现值', f'{pv_tv_2:,.2f}'],
            ['股权价值', f'{equity_2:,.2f}'],
        ]
        utils.add_table_data(document, fcfe_headers, fcfe_data)

        # Show projected FCFE details
        cf_list_2 = dcf_result.get('method2_cf_list', [])
        if cf_list_2:
            proj_headers = ['年份', 'FCFE(万元)', '折现因子', '现值(万元)']
            proj_data = []
            for cf in cf_list_2:
                proj_data.append([
                    f'第{cf["year"]}年',
                    f'{cf["fcfe"]:,.2f}',
                    f'{cf["discount_factor"]:.4f}',
                    f'{cf["pv"]:,.2f}',
                ])
            utils.add_table_data(document, proj_headers, proj_data)

        # ---- Method 3: APV (Adjusted Present Value) ----
        utils.add_paragraph(document, '')
        utils.add_paragraph(document, '方法三：APV（调整现值法）', bold=True)

        equity_3 = dcf_result.get('method3_equity_value', 0)
        net_debt = dcf_result.get('net_debt', 0)
        unlevered_value = dcf_result.get('method3_unlevered_value', 0)
        pv_tax_shield = dcf_result.get('method3_pv_tax_shield', 0)
        apv_firm_value = dcf_result.get('method3_apv_firm_value', 0)
        unlevered_beta = dcf_result.get('method3_unlevered_beta', 0)
        unlevered_ke = dcf_result.get('method3_unlevered_ke', 0)

        utils.add_paragraph(
            document,
            f'APV法将企业价值分解为无杠杆企业价值与税盾现值之和。'
            f'无杠杆Beta={unlevered_beta:.4f}，无杠杆股权成本={unlevered_ke:.4f}。'
        )

        apv_headers = ['项目', '金额(万元)']
        apv_data = [
            ['无杠杆企业价值', f'{unlevered_value:,.2f}'],
            ['加：税盾现值', f'{pv_tax_shield:,.2f}'],
            ['企业价值(APV)', f'{apv_firm_value:,.2f}'],
            ['减：净债务', f'{net_debt:,.2f}'],
            ['股权价值', f'{equity_3:,.2f}'],
        ]
        utils.add_table_data(document, apv_headers, apv_data)

        # Show APV projected FCF details (discounted at unlevered Ke)
        cf_list_3 = dcf_result.get('method3_cf_list', [])
        if cf_list_3:
            proj_headers = ['年份', 'FCF(万元)', '折现因子(Ku)', '现值(万元)']
            proj_data = []
            for cf in cf_list_3:
                proj_data.append([
                    f'第{cf["year"]}年',
                    f'{cf["fcf"]:,.2f}',
                    f'{cf["discount_factor"]:.4f}',
                    f'{cf["pv"]:,.2f}',
                ])
            utils.add_table_data(document, proj_headers, proj_data)

        # ==================== 4.4 每股价值汇总表 ====================
        utils.add_title(document, '4.4 每股价值汇总', level=2)

        summary_df = calculator.generate_valuation_summary(
            dcf_result,
            current_price=current_price,
            total_shares=total_share,
        )

        if not summary_df.empty:
            sum_headers = ['估值方法', '企业价值(万元)', '股权价值(万元)', '每股价值(元)', '安全边际(%)']
            sum_data = []
            for _, row in summary_df.iterrows():
                sum_data.append([
                    str(row.get('method', '')),
                    f'{row["enterprise_value"]:,.2f}' if isinstance(row.get('enterprise_value'), (int, float)) else str(row.get('enterprise_value', '')),
                    f'{row["equity_value"]:,.2f}',
                    f'{row["per_share"]:.2f}',
                    f'{row["margin_pct"]:.1f}%',
                ])
            utils.add_table_data(document, sum_headers, sum_data)

            # Average per-share value
            avg_per_share = summary_df['per_share'].mean()
            utils.add_paragraph(
                document,
                f'三种方法平均每股价值为{avg_per_share:.2f}元，'
                f'当前股价{current_price:.2f}元，'
                f'平均安全边际为{(avg_per_share / current_price - 1) * 100:.1f}%。'
                if current_price > 0 else
                f'三种方法平均每股价值为{avg_per_share:.2f}元。'
            )
        else:
            utils.add_paragraph(document, '无法计算每股价值（总股本数据缺失）。')

        # ==================== 4.5 敏感性矩阵热力图 ====================
        utils.add_title(document, '4.5 WACC×增长率敏感性分析', level=2)

        if current_price > 0 and total_share > 0:
            sensitivity = calculator.generate_sensitivity_matrix(
                dcf_result,
                current_price=current_price,
                total_shares=total_share,
                wacc_range=(0.06, 0.14, 0.01),
                growth_range=(0.00, 0.18, 0.02),
            )

            if sensitivity and sensitivity.get('matrix'):
                wacc_vals = sensitivity['wacc_values']
                growth_vals = sensitivity['growth_values']
                matrix = sensitivity['matrix']

                # Build data matrix for heatmap (per-share values)
                heatmap_data = []
                row_labels = []
                for i, w in enumerate(wacc_vals):
                    row_labels.append(f'WACC={w * 100:.1f}%')
                    row_data = []
                    for j, g in enumerate(growth_vals):
                        if i < len(matrix) and j < len(matrix[i]):
                            row_data.append(matrix[i][j]['per_share'])
                        else:
                            row_data.append(0)
                    heatmap_data.append(row_data)

                col_labels = [f'g={g * 100:.0f}%' for g in growth_vals]

                chart_path = os.path.join(IMAGES_DIR, f'{stock_code}_sensitivity_heatmap.png')
                center_val = current_price

                utils.generate_heatmap(
                    heatmap_data,
                    row_labels,
                    col_labels,
                    chart_path,
                    title=f'{stock_name} DCF敏感性分析（每股价值/元）',
                    center=center_val,
                    fmt='.1f',
                )
                utils.add_image(document, chart_path, width=utils.Inches(5.5))
                utils.add_paragraph(document, f'图4-1 {stock_name} WACC×增长率敏感性矩阵')

                utils.add_paragraph(
                    document,
                    f'注：红色区域表示估值高于当前股价{current_price:.2f}元（低估），'
                    f'绿色区域表示估值低于当前股价（高估）。'
                )
            else:
                utils.add_paragraph(document, '敏感性矩阵生成失败。')
        else:
            utils.add_paragraph(document, '缺少股价或总股本数据，无法生成敏感性分析。')

        # Store DCF results in context for cross-chapter use
        context.setdefault('results', {})
        context['results']['dcf_result'] = dcf_result
        context['results']['dcf_summary'] = summary_df.to_dict('records') if not summary_df.empty else []
        context['results']['dcf_avg_per_share'] = summary_df['per_share'].mean() if not summary_df.empty else 0
        context['results']['dcf_safety_margin'] = (
            (summary_df['per_share'].mean() / current_price - 1) * 100
            if not summary_df.empty and current_price > 0 else 0
        )

    utils.add_section_break(document)

    return context
