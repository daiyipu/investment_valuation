#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第二章：财务分析
利润表摘要、资产负债表摘要、现金流量表摘要、关键财务比率、营收与FCF趋势图
"""

import os
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

import module_utils as utils


def _safe_val(row, field):
    """Safely extract a numeric value from a DataFrame row."""
    v = row.get(field)
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    return float(v)


def _fmt_wan(value):
    """Format value to wan yuan (万元) string."""
    if value is None:
        return '-'
    return f'{value / 10000:,.2f}'


def _fmt_pct(value):
    """Format value as percentage string."""
    if value is None:
        return '-'
    return f'{value:.2f}%'


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


def _period_label(period_str):
    """Convert period string like '20231231' to '2023年'."""
    s = str(period_str)
    return f'{s[:4]}'


def generate_chapter(context):
    """第二章：财务分析"""
    document = context['document']
    stock_code = context['stock_code']
    stock_name = context['stock_name']
    font_prop = context['font_prop']
    financial_statements = context.get('financial_statements', {})
    financial_indicators = context.get('financial_indicators', pd.DataFrame())
    IMAGES_DIR = context['IMAGES_DIR']

    income_df = financial_statements.get('income', pd.DataFrame())
    bs_df = financial_statements.get('balancesheet', pd.DataFrame())
    cashflow_df = financial_statements.get('cashflow', pd.DataFrame())

    utils.add_title(document, '第二章 财务分析', level=1)

    # ==================== 2.1 利润表摘要 ====================
    utils.add_title(document, '2.1 利润表摘要（近5年）', level=2)

    if not income_df.empty:
        periods = _get_annual_periods(income_df, 5)
        if periods:
            headers = ['指标'] + [_period_label(p) for p in periods]
            data_rows = []

            # 营业收入 (total_revenue or total_rev)
            row = ['营业收入(万元)']
            for p in periods:
                r = _get_row_for_period(income_df, p)
                val = _safe_val(r, 'total_revenue') if r is not None else None
                if val is None and r is not None:
                    val = _safe_val(r, 'total_rev')
                row.append(_fmt_wan(val))
            data_rows.append(row)

            # 营业成本 (oper_cost or total_cogs)
            row = ['营业成本(万元)']
            cogs_row = ['营业总成本(万元)']
            for p in periods:
                r = _get_row_for_period(income_df, p)
                val = _safe_val(r, 'oper_cost') if r is not None else None
                row.append(_fmt_wan(val))
                val_total = _safe_val(r, 'total_cogs') if r is not None else None
                cogs_row.append(_fmt_wan(val_total))
            data_rows.append(row)
            data_rows.append(cogs_row)

            # 净利润
            row = ['净利润(万元)']
            for p in periods:
                r = _get_row_for_period(income_df, p)
                val = _safe_val(r, 'n_income') if r is not None else None
                row.append(_fmt_wan(val))
            data_rows.append(row)

            # 毛利率 = (营业收入 - 营业成本) / 营业收入
            row = ['毛利率']
            for p in periods:
                r = _get_row_for_period(income_df, p)
                rev = _safe_val(r, 'total_revenue') if r is not None else None
                if rev is None and r is not None:
                    rev = _safe_val(r, 'total_rev')
                cost = _safe_val(r, 'oper_cost') if r is not None else None
                if rev and cost and rev != 0:
                    row.append(_fmt_pct((rev - cost) / rev * 100))
                else:
                    row.append('-')
            data_rows.append(row)

            # 净利率 = 净利润 / 营业收入
            row = ['净利率']
            for p in periods:
                r = _get_row_for_period(income_df, p)
                rev = _safe_val(r, 'total_revenue') if r is not None else None
                if rev is None and r is not None:
                    rev = _safe_val(r, 'total_rev')
                ni = _safe_val(r, 'n_income') if r is not None else None
                if rev and ni and rev != 0:
                    row.append(_fmt_pct(ni / rev * 100))
                else:
                    row.append('-')
            data_rows.append(row)

            utils.add_table_data(document, headers, data_rows)
            utils.add_paragraph(document, '单位：万元')
        else:
            utils.add_paragraph(document, '暂无年度利润表数据。')
    else:
        utils.add_paragraph(document, '暂无利润表数据。')

    # ==================== 2.2 资产负债表摘要 ====================
    utils.add_title(document, '2.2 资产负债表摘要（近5年）', level=2)

    if not bs_df.empty:
        periods = _get_annual_periods(bs_df, 5)
        if periods:
            headers = ['指标'] + [_period_label(p) for p in periods]
            data_rows = []

            # 总资产
            row = ['总资产(万元)']
            for p in periods:
                r = _get_row_for_period(bs_df, p)
                val = _safe_val(r, 'total_assets') if r is not None else None
                row.append(_fmt_wan(val))
            data_rows.append(row)

            # 总负债
            row = ['总负债(万元)']
            for p in periods:
                r = _get_row_for_period(bs_df, p)
                val = _safe_val(r, 'total_liab') if r is not None else None
                row.append(_fmt_wan(val))
            data_rows.append(row)

            # 净资产（归属母公司股东权益）
            row = ['净资产(万元)']
            for p in periods:
                r = _get_row_for_period(bs_df, p)
                val = _safe_val(r, 'total_hlder_eqy_exc_min_int') if r is not None else None
                row.append(_fmt_wan(val))
            data_rows.append(row)

            # 资产负债率
            row = ['资产负债率']
            for p in periods:
                r = _get_row_for_period(bs_df, p)
                total_assets = _safe_val(r, 'total_assets') if r is not None else None
                total_liab = _safe_val(r, 'total_liab') if r is not None else None
                if total_assets and total_liab and total_assets != 0:
                    row.append(_fmt_pct(total_liab / total_assets * 100))
                else:
                    row.append('-')
            data_rows.append(row)

            utils.add_table_data(document, headers, data_rows)
            utils.add_paragraph(document, '单位：万元')
        else:
            utils.add_paragraph(document, '暂无年度资产负债表数据。')
    else:
        utils.add_paragraph(document, '暂无资产负债表数据。')

    # ==================== 2.3 现金流量表摘要 ====================
    utils.add_title(document, '2.3 现金流量表摘要（近5年）', level=2)

    if not cashflow_df.empty:
        periods = _get_annual_periods(cashflow_df, 5)
        if periods:
            headers = ['指标'] + [_period_label(p) for p in periods]
            data_rows = []

            # 经营活动现金流
            row = ['经营活动现金流(万元)']
            for p in periods:
                r = _get_row_for_period(cashflow_df, p)
                val = _safe_val(r, 'n_cashflow_act') if r is not None else None
                row.append(_fmt_wan(val))
            data_rows.append(row)

            # 投资活动现金流
            row = ['投资活动现金流(万元)']
            for p in periods:
                r = _get_row_for_period(cashflow_df, p)
                val = _safe_val(r, 'n_cashflow_inv_act') if r is not None else None
                row.append(_fmt_wan(val))
            data_rows.append(row)

            # 筹资活动现金流
            row = ['筹资活动现金流(万元)']
            for p in periods:
                r = _get_row_for_period(cashflow_df, p)
                val = _safe_val(r, 'n_cashflow_fnc_act') if r is not None else None
                row.append(_fmt_wan(val))
            data_rows.append(row)

            utils.add_table_data(document, headers, data_rows)
            utils.add_paragraph(document, '单位：万元')
        else:
            utils.add_paragraph(document, '暂无年度现金流量表数据。')
    else:
        utils.add_paragraph(document, '暂无现金流量表数据。')

    # ==================== 2.4 关键财务比率 ====================
    utils.add_title(document, '2.4 关键财务比率', level=2)

    # First try to use financial_indicators DataFrame from Tushare fina_indicator
    if not financial_indicators.empty:
        fi = financial_indicators.copy()
        fi['end_date_str'] = fi['end_date'].astype(str)
        fi_periods = _get_annual_periods(fi, 5)
        if fi_periods:
            headers = ['指标'] + [_period_label(p) for p in fi_periods]
            data_rows = []

            # ROE (weighted average ROE)
            row = ['ROE(%)']
            for p in fi_periods:
                r = _get_row_for_period(fi, p)
                val = _safe_val(r, 'roe') if r is not None else None
                if val is None and r is not None:
                    val = _safe_val(r, 'roe_waa')
                row.append(f'{val:.2f}' if val is not None else '-')
            data_rows.append(row)

            # ROA
            row = ['ROA(%)']
            for p in fi_periods:
                r = _get_row_for_period(fi, p)
                val = _safe_val(r, 'roa') if r is not None else None
                row.append(f'{val:.2f}' if val is not None else '-')
            data_rows.append(row)

            # Current ratio
            row = ['流动比率']
            for p in fi_periods:
                r = _get_row_for_period(fi, p)
                val = _safe_val(r, 'current_ratio') if r is not None else None
                row.append(f'{val:.2f}' if val is not None else '-')
            data_rows.append(row)

            # Quick ratio
            row = ['速动比率']
            for p in fi_periods:
                r = _get_row_for_period(fi, p)
                val = _safe_val(r, 'quick_ratio') if r is not None else None
                if val is None and r is not None:
                    val = _safe_val(r, 'qck_ratio')
                row.append(f'{val:.2f}' if val is not None else '-')
            data_rows.append(row)

            utils.add_table_data(document, headers, data_rows)
        else:
            utils.add_paragraph(document, '暂无年度财务指标数据。')
    else:
        # Compute ratios manually from financial statements
        utils.add_paragraph(document, '基于财务报表计算的财务比率：')

        inc_periods = _get_annual_periods(income_df, 5)
        bs_periods = _get_annual_periods(bs_df, 5)
        cf_periods = _get_annual_periods(cashflow_df, 5)

        if inc_periods and bs_periods:
            # Use the intersection of periods available
            common_periods = sorted(set(inc_periods) & set(bs_periods), reverse=True)[:5]
            if common_periods:
                headers = ['指标'] + [_period_label(p) for p in common_periods]
                data_rows = []

                # ROE = 净利润 / 净资产
                row = ['ROE(%)']
                for p in common_periods:
                    r_inc = _get_row_for_period(income_df, p)
                    r_bs = _get_row_for_period(bs_df, p)
                    ni = _safe_val(r_inc, 'n_income') if r_inc is not None else None
                    equity = _safe_val(r_bs, 'total_hlder_eqy_exc_min_int') if r_bs is not None else None
                    if ni and equity and equity != 0:
                        row.append(f'{ni / equity * 100:.2f}')
                    else:
                        row.append('-')
                data_rows.append(row)

                # ROA = 净利润 / 总资产
                row = ['ROA(%)']
                for p in common_periods:
                    r_inc = _get_row_for_period(income_df, p)
                    r_bs = _get_row_for_period(bs_df, p)
                    ni = _safe_val(r_inc, 'n_income') if r_inc is not None else None
                    total_assets = _safe_val(r_bs, 'total_assets') if r_bs is not None else None
                    if ni and total_assets and total_assets != 0:
                        row.append(f'{ni / total_assets * 100:.2f}')
                    else:
                        row.append('-')
                data_rows.append(row)

                # Current ratio = 流动资产 / 流动负债
                row = ['流动比率']
                for p in common_periods:
                    r_bs = _get_row_for_period(bs_df, p)
                    cur_assets = _safe_val(r_bs, 'total_cur_assets') if r_bs is not None else None
                    cur_liab = _safe_val(r_bs, 'total_cur_liab') if r_bs is not None else None
                    if cur_assets and cur_liab and cur_liab != 0:
                        row.append(f'{cur_assets / cur_liab:.2f}')
                    else:
                        row.append('-')
                data_rows.append(row)

                # Quick ratio = (流动资产 - 存货) / 流动负债
                row = ['速动比率']
                for p in common_periods:
                    r_bs = _get_row_for_period(bs_df, p)
                    cur_assets = _safe_val(r_bs, 'total_cur_assets') if r_bs is not None else None
                    inventories = _safe_val(r_bs, 'inventories') if r_bs is not None else None
                    cur_liab = _safe_val(r_bs, 'total_cur_liab') if r_bs is not None else None
                    if cur_assets and cur_liab and cur_liab != 0:
                        inv_val = inventories if inventories else 0
                        row.append(f'{(cur_assets - inv_val) / cur_liab:.2f}')
                    else:
                        row.append('-')
                data_rows.append(row)

                utils.add_table_data(document, headers, data_rows)
            else:
                utils.add_paragraph(document, '财务数据期间不足，无法计算财务比率。')
        else:
            utils.add_paragraph(document, '暂无足够财务报表数据用于计算比率。')

    # ==================== 2.5 营收与FCF趋势图 ====================
    utils.add_title(document, '2.5 营收与FCF趋势', level=2)

    # Build chart data
    has_chart_data = False
    if not income_df.empty and not cashflow_df.empty:
        inc_periods = _get_annual_periods(income_df, 5)
        cf_periods = _get_annual_periods(cashflow_df, 5)
        common_periods = sorted(set(inc_periods) & set(cf_periods))

        if len(common_periods) >= 2:
            has_chart_data = True
            years = [_period_label(p) for p in common_periods]
            revenues = []
            fcfs = []

            for p in common_periods:
                # Revenue
                r_inc = _get_row_for_period(income_df, p)
                rev = _safe_val(r_inc, 'total_revenue') if r_inc is not None else None
                if rev is None and r_inc is not None:
                    rev = _safe_val(r_inc, 'total_rev')
                revenues.append(rev / 10000 if rev else 0)

                # FCF = Operating CF - CapEx
                r_cf = _get_row_for_period(cashflow_df, p)
                ocf = _safe_val(r_cf, 'n_cashflow_act') if r_cf is not None else 0
                capex = _safe_val(r_cf, 'c_pay_acq_const_fiolta') if r_cf is not None else None
                if capex is not None and capex != 0:
                    fcf = (ocf - capex) / 10000
                else:
                    inv_cf = _safe_val(r_cf, 'n_cashflow_inv_act') if r_cf is not None else 0
                    fcf = (ocf - abs(inv_cf)) / 10000 if inv_cf else ocf * 0.7 / 10000
                fcfs.append(fcf)

            # Generate chart using module_utils generate_financial_trend_chart
            chart_path = os.path.join(IMAGES_DIR, f'{stock_code}_revenue_fcf_trend.png')

            fig, ax = plt.subplots(figsize=(12, 6))
            ax.bar(years, revenues, alpha=0.5, color='steelblue', label='营业收入')
            ax2 = ax.twinx()
            ax2.plot(years, fcfs, 'ro-', linewidth=2, markersize=8, label='自由现金流(FCF)')
            ax.set_xlabel('年度', fontproperties=font_prop, fontsize=12)
            ax.set_ylabel('营业收入(万元)', fontproperties=font_prop, fontsize=12)
            ax2.set_ylabel('FCF(万元)', fontproperties=font_prop, fontsize=12)
            ax.set_title(f'{stock_name}营收与FCF趋势', fontproperties=font_prop, fontsize=14, fontweight='bold')
            # Combined legend
            lines1, labels1 = ax.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            ax.legend(lines1 + lines2, labels1 + labels2, prop=font_prop, fontsize=10, loc='upper left')
            ax.grid(True, alpha=0.3)
            for label in ax.get_xticklabels():
                label.set_fontproperties(font_prop)
            for label in ax.get_yticklabels():
                label.set_fontproperties(font_prop)
            for label in ax2.get_yticklabels():
                label.set_fontproperties(font_prop)
            plt.tight_layout()
            plt.savefig(chart_path, dpi=150, bbox_inches='tight')
            plt.close()

            utils.add_image(document, chart_path, width=utils.Inches(5.5))
            utils.add_paragraph(document, f'图2-1 {stock_name}营收与FCF趋势')
        else:
            has_chart_data = False

    if not has_chart_data:
        utils.add_paragraph(document, '暂无足够数据生成营收与FCF趋势图。')

    utils.add_section_break(document)

    return context
