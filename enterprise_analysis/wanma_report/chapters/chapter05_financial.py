#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第五章：公司财务情况
包含：合并三大报表展示、财务报表分析、财务指标评分、行业对比
"""

from typing import Dict, Any, List, Optional
import pandas as pd
import numpy as np


class Chapter05Financial:
    """公司财务情况章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any], scoring_engine):
        self.config = config
        self.data = data
        self.scoring_engine = scoring_engine
        self.project_info = config.get('project', {})

    def generate(self) -> List[Dict[str, Any]]:
        elements = []

        elements.append({
            'type': 'heading',
            'content': '第五章 公司财务情况',
            'level': 1
        })

        # 5.1 合并资产负债表
        elements.append({
            'type': 'heading',
            'content': '5.1 合并资产负债表',
            'level': 2
        })
        elements.extend(self._generate_balance_sheet())

        # 5.2 合并利润表
        elements.append({
            'type': 'heading',
            'content': '5.2 合并利润表',
            'level': 2
        })
        elements.extend(self._generate_income_statement())

        # 5.3 合并现金流量表
        elements.append({
            'type': 'heading',
            'content': '5.3 合并现金流量表',
            'level': 2
        })
        elements.extend(self._generate_cashflow_statement())

        # 5.4 财务报表分析
        elements.append({
            'type': 'heading',
            'content': '5.4 财务报表分析',
            'level': 2
        })
        elements.extend(self._generate_statement_analysis())

        # 5.5 财务指标评分
        elements.append({
            'type': 'heading',
            'content': '5.5 财务指标评分',
            'level': 2
        })
        elements.extend(self._generate_financial_score())

        # 5.6 行业对比分析
        elements.append({
            'type': 'heading',
            'content': '5.6 行业对比分析',
            'level': 2
        })
        elements.extend(self._generate_industry_comparison())

        # 5.7 分红情况分析
        elements.append({
            'type': 'heading',
            'content': '5.7 分红情况分析',
            'level': 2
        })
        elements.extend(self._generate_dividend_analysis())

        return elements

    # ============================================================
    # 5.1 合并资产负债表
    # ============================================================
    def _generate_balance_sheet(self) -> List[Dict[str, Any]]:
        elements = []
        bs = self.data.get('financial_statements', {}).get('balance_sheet', pd.DataFrame())

        if bs.empty:
            elements.append({'type': 'paragraph', 'content': '暂无资产负债表数据。'})
            return elements

        periods = self._get_annual_periods(bs, 3)
        if not periods:
            elements.append({'type': 'paragraph', 'content': '暂无年度资产负债表数据。'})
            return elements

        # 定义要展示的科目和显示名称
        asset_items = [
            ('money_cap', '货币资金'),
            ('account_rece', '应收账款'),
            ('prepayment', '预付款项'),
            ('invty', '存货'),
            ('oth_cur_assets', '其他流动资产'),
            ('total_cur_assets', '流动资产合计'),
            ('fix_assets', '固定资产'),
            ('intan_assets', '无形资产'),
            ('total_assets', '资产总计'),
        ]
        liability_items = [
            ('st_loan', '短期借款'),
            ('acct_payable', '应付账款'),
            ('oth_payable', '其他应付款'),
            ('total_cur_liab', '流动负债合计'),
            ('lt_borr', '长期借款'),
            ('total_liab', '负债合计'),
        ]
        equity_items = [
            ('total_hldr_eqy_exc_min_int', '股东权益合计'),
        ]

        # 构建表格
        rows = []
        for items, section_name in [(asset_items, '资产'), (liability_items, '负债'), (equity_items, '权益')]:
            for field, name in items:
                row = {'科目': name}
                for period in periods:
                    val = self._get_value_for_period(bs, field, period)
                    row[self._period_label(period)] = self._format_wan(val)
                rows.append(row)
            # 添加空行分隔
            if section_name != '权益':
                row = {'科目': ''}
                for period in periods:
                    row[self._period_label(period)] = ''
                rows.append(row)

        headers = ['科目'] + [self._period_label(p) for p in periods]
        elements.append({
            'type': 'table',
            'headers': headers,
            'data': rows
        })
        elements.append({'type': 'paragraph', 'content': '单位：万元'})

        return elements

    # ============================================================
    # 5.2 合并利润表
    # ============================================================
    def _generate_income_statement(self) -> List[Dict[str, Any]]:
        elements = []
        ic = self.data.get('financial_statements', {}).get('income_statement', pd.DataFrame())

        if ic.empty:
            elements.append({'type': 'paragraph', 'content': '暂无利润表数据。'})
            return elements

        periods = self._get_annual_periods(ic, 3)
        if not periods:
            elements.append({'type': 'paragraph', 'content': '暂无年度利润表数据。'})
            return elements

        items = [
            ('total_revenue', '营业总收入'),
            ('total_cogs', '营业总成本'),
            ('operate_profit', '营业利润'),
            ('total_profit', '利润总额'),
            ('n_income', '净利润'),
            ('n_income_attr_p', '归属于母公司所有者的净利润'),
            ('rd_exp', '研发费用'),
            ('sell_exp', '销售费用'),
            ('admin_exp', '管理费用'),
            ('fin_exp', '财务费用'),
        ]

        rows = []
        for field, name in items:
            row = {'科目': name}
            for period in periods:
                val = self._get_value_for_period(ic, field, period)
                row[self._period_label(period)] = self._format_wan(val)
            rows.append(row)

        headers = ['科目'] + [self._period_label(p) for p in periods]
        elements.append({
            'type': 'table',
            'headers': headers,
            'data': rows
        })
        elements.append({'type': 'paragraph', 'content': '单位：万元'})

        return elements

    # ============================================================
    # 5.3 合并现金流量表
    # ============================================================
    def _generate_cashflow_statement(self) -> List[Dict[str, Any]]:
        elements = []
        cf = self.data.get('financial_statements', {}).get('cashflow', pd.DataFrame())

        if cf.empty:
            elements.append({'type': 'paragraph', 'content': '暂无现金流量表数据。'})
            return elements

        periods = self._get_annual_periods(cf, 3)
        if not periods:
            elements.append({'type': 'paragraph', 'content': '暂无年度现金流量表数据。'})
            return elements

        items = [
            ('n_cashflow_act', '经营活动现金流净额'),
            ('n_cashflow_inv_act', '投资活动现金流净额'),
            ('n_cashflow_fnc_act', '筹资活动现金流净额'),
            ('c_cashflow_end_bal', '期末现金及现金等价物余额'),
        ]

        rows = []
        for field, name in items:
            row = {'科目': name}
            for period in periods:
                val = self._get_value_for_period(cf, field, period)
                row[self._period_label(period)] = self._format_wan(val)
            rows.append(row)

        headers = ['科目'] + [self._period_label(p) for p in periods]
        elements.append({
            'type': 'table',
            'headers': headers,
            'data': rows
        })
        elements.append({'type': 'paragraph', 'content': '单位：万元'})

        return elements

    # ============================================================
    # 5.4 财务报表分析
    # ============================================================
    def _generate_statement_analysis(self) -> List[Dict[str, Any]]:
        elements = []
        bs = self.data.get('financial_statements', {}).get('balance_sheet', pd.DataFrame())
        ic = self.data.get('financial_statements', {}).get('income_statement', pd.DataFrame())
        cf = self.data.get('financial_statements', {}).get('cashflow', pd.DataFrame())

        # 资产结构分析
        elements.append({
            'type': 'heading',
            'content': '5.4.1 资产结构分析',
            'level': 3
        })
        elements.extend(self._analyze_asset_structure(bs))

        # 负债结构分析
        elements.append({
            'type': 'heading',
            'content': '5.4.2 负债结构分析',
            'level': 3
        })
        elements.extend(self._analyze_liability_structure(bs))

        # 盈利能力分析
        elements.append({
            'type': 'heading',
            'content': '5.4.3 盈利能力分析',
            'level': 3
        })
        elements.extend(self._analyze_profitability(ic))

        # 现金流分析
        elements.append({
            'type': 'heading',
            'content': '5.4.4 现金流分析',
            'level': 3
        })
        elements.extend(self._analyze_cashflow(cf))

        return elements

    def _analyze_asset_structure(self, bs: pd.DataFrame) -> List[Dict[str, Any]]:
        elements = []
        if bs.empty:
            elements.append({'type': 'paragraph', 'content': '暂无资产数据。'})
            return elements

        periods = self._get_annual_periods(bs, 2)
        if len(periods) < 1:
            elements.append({'type': 'paragraph', 'content': '数据不足。'})
            return elements

        latest = self._get_row_for_period(bs, periods[0])
        total_assets = self._val(latest, 'total_assets')

        if total_assets and total_assets > 0:
            cur_assets = self._val(latest, 'total_cur_assets') or 0
            cur_pct = cur_assets / total_assets * 100

            invty = self._val(latest, 'invty') or 0
            ar = self._val(latest, 'account_rece') or 0
            money = self._val(latest, 'money_cap') or 0

            text = (
                f"截至最新报告期，公司总资产{total_assets / 10000:,.2f}万元，"
                f"其中流动资产{cur_assets / 10000:,.2f}万元（占比{cur_pct:.1f}%）。"
                f"主要资产科目：货币资金{money / 10000:,.2f}万元、"
                f"应收账款{ar / 10000:,.2f}万元、存货{invty / 10000:,.2f}万元。"
            )
            elements.append({'type': 'paragraph', 'content': text})

            # 同比变动（如有两年数据）
            if len(periods) >= 2:
                prev = self._get_row_for_period(bs, periods[1])
                prev_total = self._val(prev, 'total_assets')
                if prev_total and prev_total > 0:
                    yoy = (total_assets - prev_total) / prev_total * 100
                    direction = '增长' if yoy > 0 else '下降'
                    elements.append({
                        'type': 'paragraph',
                        'content': f"总资产同比{direction}{abs(yoy):.2f}%。"
                    })

        return elements

    def _analyze_liability_structure(self, bs: pd.DataFrame) -> List[Dict[str, Any]]:
        elements = []
        if bs.empty:
            return elements

        periods = self._get_annual_periods(bs, 2)
        if not periods:
            return elements

        latest = self._get_row_for_period(bs, periods[0])
        total_liab = self._val(latest, 'total_liab')
        total_assets = self._val(latest, 'total_assets')

        if total_liab and total_assets and total_assets > 0:
            debt_ratio = total_liab / total_assets * 100

            st_loan = self._val(latest, 'st_loan') or 0
            lt_borr = self._val(latest, 'lt_borr') or 0
            interest_debt = st_loan + lt_borr

            operating_liab = total_liab - interest_debt
            operating_pct = operating_liab / total_liab * 100 if total_liab > 0 else 0

            text = (
                f"公司资产负债率{debt_ratio:.2f}%。"
                f"有息负债（短期借款+长期借款）{interest_debt / 10000:,.2f}万元，"
                f"经营性负债占比{operating_pct:.1f}%。"
            )
            elements.append({'type': 'paragraph', 'content': text})

        return elements

    def _analyze_profitability(self, ic: pd.DataFrame) -> List[Dict[str, Any]]:
        elements = []
        if ic.empty:
            return elements

        periods = self._get_annual_periods(ic, 3)
        if not periods:
            return elements

        # 收入和利润趋势表
        rows = []
        for period in periods:
            row_data = self._get_row_for_period(ic, period)
            revenue = self._val(row_data, 'total_revenue')
            net_income = self._val(row_data, 'n_income_attr_p') or self._val(row_data, 'n_income')
            rd_exp = self._val(row_data, 'rd_exp')

            row = {
                '年度': self._period_label(period),
                '营业总收入(万元)': self._format_wan(revenue),
                '净利润(万元)': self._format_wan(net_income),
                '研发费用(万元)': self._format_wan(rd_exp),
            }
            if revenue and net_income:
                row['净利率'] = f"{net_income / revenue * 100:.2f}%"
            rows.append(row)

        if rows:
            elements.append({
                'type': 'table',
                'headers': list(rows[0].keys()),
                'data': rows
            })

        # 文字分析
        if len(periods) >= 2:
            latest_row = self._get_row_for_period(ic, periods[0])
            prev_row = self._get_row_for_period(ic, periods[1])
            rev = self._val(latest_row, 'total_revenue')
            prev_rev = self._val(prev_row, 'total_revenue')
            if rev and prev_rev and prev_rev > 0:
                yoy = (rev - prev_rev) / prev_rev * 100
                direction = '增长' if yoy > 0 else '下降'
                elements.append({
                    'type': 'paragraph',
                    'content': f"营业收入同比{direction}{abs(yoy):.2f}%。"
                })

        return elements

    def _analyze_cashflow(self, cf: pd.DataFrame) -> List[Dict[str, Any]]:
        elements = []
        if cf.empty:
            return elements

        periods = self._get_annual_periods(cf, 3)
        if not periods:
            return elements

        rows = []
        for period in periods:
            row_data = self._get_row_for_period(cf, period)
            oper = self._val(row_data, 'n_cashflow_act')
            invest = self._val(row_data, 'n_cashflow_inv_act')
            finance = self._val(row_data, 'n_cashflow_fnc_act')

            rows.append({
                '年度': self._period_label(period),
                '经营活动(万元)': self._format_wan(oper),
                '投资活动(万元)': self._format_wan(invest),
                '筹资活动(万元)': self._format_wan(finance),
            })

        if rows:
            elements.append({
                'type': 'table',
                'headers': list(rows[0].keys()),
                'data': rows
            })

        # 现金流特征分析
        latest_row = self._get_row_for_period(cf, periods[0])
        oper = self._val(latest_row, 'n_cashflow_act')
        if oper is not None:
            status = '为正' if oper > 0 else '为负'
            text = f"最新报告期经营活动现金流{status}，金额{oper / 10000:,.2f}万元。"
            if oper < 0:
                text += "需关注公司经营造血能力。"
            elements.append({'type': 'paragraph', 'content': text})

        return elements

    # ============================================================
    # 5.5 财务指标评分
    # ============================================================
    def _generate_financial_score(self) -> List[Dict[str, Any]]:
        elements = []
        scoring_result = self.data.get('scoring_result', {})

        dim_name_map = {
            'operations': '运营能力',
            'profitability': '盈利能力',
            'solvency': '偿债能力',
            'growth': '成长能力',
        }

        # 综合评分表
        dim_table = []
        for dim_key, dim_name in dim_name_map.items():
            dim_data = scoring_result.get('dimensions', {}).get(dim_key, {})
            dim_table.append({
                '维度': dim_name,
                '权重': f"{dim_data.get('weight', 0) * 100:.0f}%",
                '得分': f"{dim_data.get('score', 0):.2f}",
                '加权得分': f"{dim_data.get('weighted_score', 0):.2f}",
            })

        elements.append({
            'type': 'table',
            'headers': ['维度', '权重', '得分', '加权得分'],
            'data': dim_table
        })

        total = scoring_result.get('total_score', 0)
        grade = scoring_result.get('grade', '暂无数据')
        elements.append({
            'type': 'paragraph',
            'content': f"综合财务评分{total:.2f}分，评级为\"{grade}\"。"
        })

        # 各维度指标详情
        for dim_key, dim_name in dim_name_map.items():
            indicators = scoring_result.get('indicators', {}).get(dim_key, [])
            if not indicators:
                continue

            # 偿债能力有子维度，特殊处理
            if dim_key == 'solvency':
                sub_dims = scoring_result.get('sub_dimensions', {}).get('solvency', {})
                for sub_key in ['short_term', 'long_term']:
                    sub = sub_dims.get(sub_key, {})
                    sub_name = sub.get('name', '')
                    sub_indicators = sub.get('indicators', [])
                    if sub_indicators:
                        elements.append({
                            'type': 'paragraph',
                            'content': f"{sub_name}（得分：{sub.get('score', 0):.1f}）："
                        })
                        elements.extend(self._format_indicator_table(sub_indicators))
                continue

            elements.append({
                'type': 'paragraph',
                'content': f"{dim_name}："
            })
            elements.extend(self._format_indicator_table(indicators))

        return elements

    def _format_indicator_table(self, indicators: list) -> List[Dict[str, Any]]:
        elements = []
        table_data = []
        for ind in indicators:
            value_str = f"{ind.get('value'):.4f}" if ind.get('value') is not None else '-'
            score_str = f"{ind.get('score'):.1f}" if ind.get('score') is not None else '-'
            table_data.append({
                '指标': ind.get('name', ''),
                '数值': value_str,
                '等级': ind.get('quantile_level', '-'),
                '得分': score_str,
            })

        if table_data:
            elements.append({
                'type': 'table',
                'headers': ['指标', '数值', '等级', '得分'],
                'data': table_data
            })
        return elements

    # ============================================================
    # 5.6 行业对比分析
    # ============================================================
    def _generate_industry_comparison(self) -> List[Dict[str, Any]]:
        elements = []
        industry_data = self.data.get('industry_data', pd.DataFrame())

        if industry_data.empty:
            elements.append({'type': 'paragraph', 'content': '暂无行业对比数据。'})
            return elements

        field_names = {
            'roe': '净资产收益率',
            'grossprofit_margin': '毛利率',
            'netprofit_margin': '净利率',
            'current_ratio': '流动比率',
            'debt_to_assets': '资产负债率',
            'assets_turn': '总资产周转率',
            'netprofit_yoy': '净利润增长率',
        }
        alias_map = {
            'grossprofit_margin': 'gross_margin',
            'netprofit_margin': 'net_margin',
            'assets_turn': 'asset_turn',
            'netprofit_yoy': 'net_profit_yoy',
        }
        key_fields = list(field_names.keys())

        comparison_data = []
        for field in key_fields:
            col = field
            if col not in industry_data.columns:
                col = alias_map.get(field, field)
            if col not in industry_data.columns:
                continue

            company_value = self._get_company_financial_value(field)
            industry_median = industry_data[col].median()
            ranking = self._calculate_ranking(company_value, col, industry_data)

            comparison_data.append({
                '指标': field_names.get(field, field),
                '公司值': f"{company_value:.4f}" if company_value is not None else '-',
                '行业中位数': f"{industry_median:.4f}" if pd.notna(industry_median) else '-',
                '行业排名': ranking,
            })

        if comparison_data:
            elements.append({
                'type': 'table',
                'headers': ['指标', '公司值', '行业中位数', '行业排名'],
                'data': comparison_data
            })

        return elements

    def _get_company_financial_value(self, field: str):
        fi = self.data.get('financial_indicators', pd.DataFrame())
        if fi.empty:
            return None
        col = field
        if col not in fi.columns:
            alias_map = {
                'grossprofit_margin': 'gross_margin',
                'netprofit_margin': 'net_margin',
                'assets_turn': 'asset_turn',
                'netprofit_yoy': 'net_profit_yoy',
                'ca_turn': 'current_asset_turn',
                'ar_turn': 'arturn_days',
                'or_yoy': 'revenue_yoy',
            }
            col = alias_map.get(field, field)
        if col not in fi.columns:
            return None
        df_sorted = fi.sort_values('end_date', ascending=False)
        val = df_sorted.iloc[0].get(col)
        return round(float(val), 4) if val is not None and pd.notna(val) else None

    def _calculate_ranking(self, company_value, field: str, industry_data: pd.DataFrame) -> str:
        if company_value is None or pd.isna(company_value):
            return '-'
        reverse_fields = {'debt_to_assets', 'arturn_days', 'debt_to_eqt', 'int_to_talcap'}
        is_reverse = field in reverse_fields

        if 'end_date' in industry_data.columns:
            latest_year = industry_data['end_date'].max()
            latest_data = industry_data[industry_data['end_date'] == latest_year]
        else:
            latest_data = industry_data

        if field not in latest_data.columns:
            return '-'

        if 'ts_code' in latest_data.columns:
            peer_values = latest_data.groupby('ts_code')[field].first().dropna()
        else:
            peer_values = latest_data[field].dropna()

        if len(peer_values) == 0:
            return '-'

        if is_reverse:
            sorted_values = peer_values.sort_values(ascending=True)
        else:
            sorted_values = peer_values.sort_values(ascending=False)

        rank = 1
        for val in sorted_values:
            if abs(float(val) - company_value) < 1e-10:
                break
            rank += 1
        rank = min(rank, len(sorted_values))

        return f"{rank}/{len(sorted_values)}"

    # ============================================================
    # 5.7 分红情况分析
    # ============================================================
    def _generate_dividend_analysis(self) -> List[Dict[str, Any]]:
        elements = []
        dividend = self.data.get('dividend', pd.DataFrame())

        if dividend.empty:
            elements.append({'type': 'paragraph', 'content': '暂无分红数据。'})
            return elements

        # 按年度分组，取最近5年
        if 'end_date' not in dividend.columns:
            elements.append({'type': 'paragraph', 'content': '分红数据格式异常。'})
            return elements

        dividend = dividend.copy()
        dividend['year'] = dividend['end_date'].astype(str).str[:4]
        years = sorted(dividend['year'].unique(), reverse=True)[:5]

        table_data = []
        for year in years:
            year_data = dividend[dividend['year'] == year]

            # 查找含税派息数据
            cash_div = None
            div_ratio = None
            for _, row in year_data.iterrows():
                cash_ps = row.get('cash_div', None) or row.get('cash_div_tax', None)
                if cash_ps and pd.notna(cash_ps):
                    cash_div = float(cash_ps)
                    break

            record_date = year_data.iloc[0].get('end_date', '') if not year_data.empty else ''
            imp_ann_date = year_data.iloc[0].get('imp_ann_date', '') if not year_data.empty else ''
            pay_date = year_data.iloc[0].get('pay_date', '') if not year_data.empty else ''
            div_type = year_data.iloc[0].get('div_type', '') if not year_data.empty else ''

            type_desc = {
                'CN': '现金分红',
                'SH': '送股',
                'ZZ': '转增',
            }.get(str(div_type), str(div_type))

            row = {
                '年度': f"{year}年",
                '分红类型': type_desc,
                '每股派息(元)': f"{cash_div:.4f}" if cash_div else '-',
            }
            table_data.append(row)

        if table_data:
            elements.append({
                'type': 'table',
                'headers': ['年度', '分红类型', '每股派息(元)'],
                'data': table_data
            })

            # 文字分析
            div_years = [r for r in table_data if r['每股派息(元)'] != '-']
            if div_years:
                total_years = len(table_data)
                div_count = len(div_years)
                avg_div = sum(float(r['每股派息(元)']) for r in div_years) / len(div_years)

                text = f"近{total_years}年中，公司有{div_count}年进行现金分红，"
                text += f"平均每股派息{avg_div:.4f}元。"

                if div_count == total_years:
                    text += "公司分红政策连续稳定。"
                elif div_count >= total_years * 0.6:
                    text += "公司基本保持分红政策。"
                else:
                    text += "公司分红频率较低。"

                elements.append({'type': 'paragraph', 'content': text})
            else:
                elements.append({
                    'type': 'paragraph',
                    'content': '近年内公司未进行现金分红。'
                })

        return elements

    # ============================================================
    # 工具方法
    # ============================================================
    def _get_annual_periods(self, df: pd.DataFrame, n: int = 3) -> list:
        """获取最近n个年报期（1231）"""
        if df.empty or 'end_date' not in df.columns:
            return []
        annual = df[df['end_date'].astype(str).str.contains('1231', na=False)]
        periods = sorted(annual['end_date'].unique(), reverse=True)
        return list(periods[:n])

    def _get_row_for_period(self, df: pd.DataFrame, period: str) -> Optional[pd.Series]:
        """获取某期数据行"""
        if df.empty:
            return None
        period_data = df[df['end_date'] == period]
        if period_data.empty:
            return None
        return period_data.iloc[0]

    def _get_value_for_period(self, df: pd.DataFrame, field: str, period: str) -> Optional[float]:
        row = self._get_row_for_period(df, period)
        if row is None:
            return None
        return self._val(row, field)

    def _val(self, row: Optional[pd.Series], field: str) -> Optional[float]:
        if row is None:
            return None
        v = row.get(field)
        if v is None or pd.isna(v):
            return None
        return float(v)

    def _format_wan(self, value: Optional[float]) -> str:
        """格式化为万元"""
        if value is None:
            return '-'
        return f"{value / 10000:,.2f}"

    def _period_label(self, period: str) -> str:
        s = str(period)
        return f"{s[:4]}年"
