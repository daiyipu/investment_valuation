#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第五章：公司财务情况
包含：合并三大报表展示（5年完整科目）、科目占比分析（含行业百分位评级）、
财务报表分析、财务指标评分、行业对比、分红情况、财务异常汇总
"""

from typing import Dict, Any, List, Optional
import pandas as pd
import numpy as np


class Chapter05Financial:
    """公司财务情况章节生成器"""

    # ============================================================
    # 完整科目定义（Tushare标准字段名）
    # ============================================================
    BS_ASSET_ITEMS = [
        ('money_cap', '货币资金'),
        ('notes_receiv', '应收票据'),
        ('accounts_receiv', '应收账款'),
        ('prepayment', '预付款项'),
        ('oth_receiv', '其他应收款'),
        ('inventories', '存货'),
        ('oth_cur_assets', '其他流动资产'),
        ('total_cur_assets', '流动资产合计'),
        ('lt_eqt_invest', '长期股权投资'),
        ('invest_real_estate', '投资性房地产'),
        ('fix_assets', '固定资产'),
        ('cip', '在建工程'),
        ('intan_assets', '无形资产'),
        ('goodwill', '商誉'),
        ('lt_amor_exp', '长期待摊费用'),
        ('defer_tax_assets', '递延所得税资产'),
        ('oth_nca', '其他非流动资产'),
        ('total_nca', '非流动资产合计'),
        ('total_assets', '资产总计'),
    ]

    BS_LIABILITY_ITEMS = [
        ('st_borr', '短期借款'),
        ('notes_payable', '应付票据'),
        ('acct_payable', '应付账款'),
        ('adv_receipts', '预收款项'),
        ('payroll_payable', '应付职工薪酬'),
        ('taxes_payable', '应交税费'),
        ('int_payable', '应付利息'),
        ('oth_payable', '其他应付款'),
        ('non_cur_liab_due_1y', '一年内到期非流动负债'),
        ('oth_cur_liab', '其他流动负债'),
        ('total_cur_liab', '流动负债合计'),
        ('lt_borr', '长期借款'),
        ('bond_payable', '应付债券'),
        ('lease_liab', '租赁负债'),
        ('total_ncl', '非流动负债合计'),
        ('total_liab', '负债合计'),
    ]

    BS_EQUITY_ITEMS = [
        ('cap_rese', '资本公积'),
        ('surplus_rese', '盈余公积'),
        ('undistr_porfit', '未分配利润'),
        ('total_hldr_eqy_exc_min_int', '归属母公司股东权益'),
        ('minority_int', '少数股东权益'),
    ]

    IS_ITEMS = [
        ('total_revenue', '营业总收入'),
        ('total_cogs', '营业总成本'),
        ('oper_cost', '营业成本'),
        ('biz_tax_surchg', '营业税金及附加'),
        ('sell_exp', '销售费用'),
        ('admin_exp', '管理费用'),
        ('fin_exp', '财务费用'),
        ('assets_impair_loss', '资产减值损失'),
        ('rd_exp', '研发费用'),
        ('operate_profit', '营业利润'),
        ('total_profit', '利润总额'),
        ('income_tax', '所得税费用'),
        ('n_income', '净利润'),
        ('n_income_attr_p', '归母净利润'),
        ('basic_eps', '基本每股收益'),
    ]

    CF_ITEMS = [
        ('c_fr_sale_sg', '销售商品收到现金'),
        ('recp_tax_rends', '收到税费返还'),
        ('n_cashflow_act', '经营活动净额'),
        ('c_pay_acq_const_fiolta', '购建长期资产支付现金'),
        ('n_cashflow_inv_act', '投资活动净额'),
        ('c_recp_borrow', '取得借款收到现金'),
        ('c_prepay_amt_borr', '偿还债务支付现金'),
        ('c_pay_dist_dpcp_int_exp', '分配股利偿付利息'),
        ('n_cash_flows_fnc_act', '筹资活动净额'),
        ('c_cash_equ_end_period', '期末现金余额'),
    ]

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
        elements.append({'type': 'heading', 'content': '5.1 合并资产负债表', 'level': 2})
        elements.extend(self._generate_balance_sheet())

        # 5.2 合并利润表
        elements.append({'type': 'heading', 'content': '5.2 合并利润表', 'level': 2})
        elements.extend(self._generate_income_statement())

        # 5.3 合并现金流量表
        elements.append({'type': 'heading', 'content': '5.3 合并现金流量表', 'level': 2})
        elements.extend(self._generate_cashflow_statement())

        # 5.4 财务报表分析
        elements.append({'type': 'heading', 'content': '5.4 财务报表分析', 'level': 2})
        elements.extend(self._generate_statement_analysis())

        # 5.5 财务指标评分
        elements.append({'type': 'heading', 'content': '5.5 财务指标评分', 'level': 2})
        elements.extend(self._generate_financial_score())

        # 5.6 行业对比分析
        elements.append({'type': 'heading', 'content': '5.6 行业对比分析', 'level': 2})
        elements.extend(self._generate_industry_comparison())

        # 5.7 分红情况分析
        elements.append({'type': 'heading', 'content': '5.7 分红情况分析', 'level': 2})
        elements.extend(self._generate_dividend_analysis())

        # 5.8 财务异常汇总
        elements.append({'type': 'heading', 'content': '5.8 财务异常汇总', 'level': 2})
        elements.extend(self._generate_anomaly_summary())

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

        periods = self._get_annual_periods(bs, 5)
        if not periods:
            elements.append({'type': 'paragraph', 'content': '暂无年度资产负债表数据。'})
            return elements

        # 5.1.1 资产负债表数据
        elements.append({'type': 'heading', 'content': '5.1.1 资产负债表数据', 'level': 3})

        rows = []
        for items in [self.BS_ASSET_ITEMS, self.BS_LIABILITY_ITEMS, self.BS_EQUITY_ITEMS]:
            for field, name in items:
                row = {'科目': name}
                for period in periods:
                    val = self._get_value_for_period(bs, field, period)
                    row[self._period_label(period)] = self._format_wan(val)
                rows.append(row)
            # 空行分隔
            rows.append({'科目': ''} | {self._period_label(p): '' for p in periods})

        # 移除最后一个空行
        if rows and rows[-1].get('科目', '') == '':
            rows.pop()

        headers = ['科目'] + [self._period_label(p) for p in periods]
        elements.append({'type': 'table', 'headers': headers, 'data': rows})
        elements.append({'type': 'paragraph', 'content': '单位：万元'})

        # 5.1.2 资产负债表科目占比分析
        elements.append({'type': 'heading', 'content': '5.1.2 资产负债表科目占比分析', 'level': 3})
        all_items = self.BS_ASSET_ITEMS + self.BS_LIABILITY_ITEMS + self.BS_EQUITY_ITEMS
        elements.extend(self._build_proportion_analysis(
            bs, periods, all_items, 'total_assets', '总资产', '资产负债表'))

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

        periods = self._get_annual_periods(ic, 5)
        if not periods:
            elements.append({'type': 'paragraph', 'content': '暂无年度利润表数据。'})
            return elements

        # 5.2.1 利润表数据
        elements.append({'type': 'heading', 'content': '5.2.1 利润表数据', 'level': 3})

        rows = []
        for field, name in self.IS_ITEMS:
            row = {'科目': name}
            for period in periods:
                val = self._get_value_for_period(ic, field, period)
                if name == '基本每股收益':
                    row[self._period_label(period)] = f"{val:.4f}" if val is not None else '-'
                else:
                    row[self._period_label(period)] = self._format_wan(val)
            rows.append(row)

        headers = ['科目'] + [self._period_label(p) for p in periods]
        elements.append({'type': 'table', 'headers': headers, 'data': rows})
        elements.append({'type': 'paragraph', 'content': '单位：万元（每股收益除外）'})

        # 5.2.2 利润表科目占比分析
        elements.append({'type': 'heading', 'content': '5.2.2 利润表科目占比分析', 'level': 3})
        # 利润表用营业总收入作为基准，排除营业总收入自身和基本每股收益
        is_items_for_pct = [(f, n) for f, n in self.IS_ITEMS if f not in ('total_revenue', 'basic_eps')]
        elements.extend(self._build_proportion_analysis(
            ic, periods, is_items_for_pct, 'total_revenue', '营业总收入', '利润表'))

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

        periods = self._get_annual_periods(cf, 5)
        if not periods:
            elements.append({'type': 'paragraph', 'content': '暂无年度现金流量表数据。'})
            return elements

        # 5.3.1 现金流量表数据
        elements.append({'type': 'heading', 'content': '5.3.1 现金流量表数据', 'level': 3})

        rows = []
        for field, name in self.CF_ITEMS:
            row = {'科目': name}
            for period in periods:
                val = self._get_value_for_period(cf, field, period)
                row[self._period_label(period)] = self._format_wan(val)
            rows.append(row)

        headers = ['科目'] + [self._period_label(p) for p in periods]
        elements.append({'type': 'table', 'headers': headers, 'data': rows})
        elements.append({'type': 'paragraph', 'content': '单位：万元'})

        # 5.3.2 现金流量表科目占比分析
        elements.append({'type': 'heading', 'content': '5.3.2 现金流量表科目占比分析', 'level': 3})
        # 现金流量表用经营现金流绝对值作为基准
        cf_items_for_pct = [(f, n) for f, n in self.CF_ITEMS if f != 'n_cashflow_act']
        elements.extend(self._build_proportion_analysis(
            cf, periods, cf_items_for_pct, 'n_cashflow_act', '经营活动现金流绝对值', '现金流量表',
            use_abs_denom=True))

        return elements

    # ============================================================
    # 科目占比分析（核心方法）
    # ============================================================
    def _build_proportion_analysis(self, df, periods, items, denom_field, denom_label,
                                   statement_type, use_abs_denom=False) -> List[Dict[str, Any]]:
        """生成科目占比分析表，含行业百分位评级"""
        elements = []
        peer_df = self._get_peer_df(statement_type)

        rows = []
        for field, name in items:
            row = {'科目': name}
            for period in periods:
                val = self._get_value_for_period(df, field, period)
                denom = self._get_value_for_period(df, denom_field, period)
                if val is not None and denom is not None and denom != 0:
                    d = abs(denom) if use_abs_denom else denom
                    if d != 0:
                        row[self._period_label(period)] = f"{val / d * 100:.1f}%"
                    else:
                        row[self._period_label(period)] = '-'
                else:
                    row[self._period_label(period)] = '-'

            # 计算最新一期行业百分位
            if peer_df is not None and not peer_df.empty and len(periods) > 0:
                pct = self._calc_item_percentile(df, peer_df, field, denom_field,
                                                  periods[0], use_abs_denom)
                if pct is not None:
                    row['行业百分位'] = f"{pct:.0f}%"
                    row['评级'] = self._percentile_to_rating(pct)
                else:
                    row['行业百分位'] = '-'
                    row['评级'] = '-'
            else:
                row['行业百分位'] = '-'
                row['评级'] = '-'

            rows.append(row)

        if rows:
            headers = ['科目'] + [self._period_label(p) for p in periods] + ['行业百分位', '评级']
            elements.append({
                'type': 'table',
                'headers': headers,
                'data': rows
            })
            elements.append({
                'type': 'paragraph',
                'content': f'注：占比 = 科目 / {denom_label}。行业百分位表示公司在同行业中的位置（基于最新年报）。'
            })

        return elements

    def _calc_item_percentile(self, company_df, peer_df, item_field, denom_field,
                               period, use_abs_denom=False) -> Optional[float]:
        """计算目标公司在同行中的百分位"""
        # 目标公司比值
        company_row = self._get_row_for_period(company_df, period)
        if company_row is None:
            return None
        item_val = self._val(company_row, item_field)
        denom_val = self._val(company_row, denom_field)
        if item_val is None or denom_val is None or denom_val == 0:
            return None
        d = abs(denom_val) if use_abs_denom else denom_val
        if d == 0:
            return None
        company_ratio = item_val / d

        # 同行比值
        peer_period = peer_df[peer_df['end_date'] == period] if 'end_date' in peer_df.columns else peer_df
        if peer_period.empty:
            return None

        ratios = []
        if 'ts_code' in peer_period.columns:
            for _, grp in peer_period.groupby('ts_code'):
                row = grp.iloc[0]
                pv = row.get(item_field)
                dv = row.get(denom_field)
                if pd.notna(pv) and pd.notna(dv) and float(dv) != 0:
                    dd = abs(float(dv)) if use_abs_denom else float(dv)
                    if dd != 0:
                        ratios.append(float(pv) / dd)
        else:
            for _, row in peer_period.iterrows():
                pv = row.get(item_field)
                dv = row.get(denom_field)
                if pd.notna(pv) and pd.notna(dv) and float(dv) != 0:
                    dd = abs(float(dv)) if use_abs_denom else float(dv)
                    if dd != 0:
                        ratios.append(float(pv) / dd)

        if len(ratios) < 3:
            return None

        import numpy as np
        pct = (np.array(ratios) < company_ratio).sum() / len(ratios) * 100
        return pct

    @staticmethod
    def _percentile_to_rating(percentile: float) -> str:
        if percentile >= 80:
            return '优秀'
        elif percentile >= 60:
            return '良好'
        elif percentile >= 40:
            return '中等'
        elif percentile >= 20:
            return '较差'
        else:
            return '极差'

    def _get_peer_df(self, statement_type: str) -> Optional[pd.DataFrame]:
        """获取对应报表的同行数据"""
        peer_statements = self.data.get('peer_statements', {})
        key_map = {
            '资产负债表': 'peer_balance_sheet',
            '利润表': 'peer_income_statement',
            '现金流量表': 'peer_cashflow',
        }
        df = peer_statements.get(key_map.get(statement_type, ''), pd.DataFrame())
        return df if not df.empty else None

    # ============================================================
    # 5.4 财务报表分析
    # ============================================================
    def _generate_statement_analysis(self) -> List[Dict[str, Any]]:
        elements = []
        bs = self.data.get('financial_statements', {}).get('balance_sheet', pd.DataFrame())
        ic = self.data.get('financial_statements', {}).get('income_statement', pd.DataFrame())
        cf = self.data.get('financial_statements', {}).get('cashflow', pd.DataFrame())

        elements.append({'type': 'heading', 'content': '5.4.1 资产结构分析', 'level': 3})
        elements.extend(self._analyze_asset_structure(bs))

        elements.append({'type': 'heading', 'content': '5.4.2 负债结构分析', 'level': 3})
        elements.extend(self._analyze_liability_structure(bs))

        elements.append({'type': 'heading', 'content': '5.4.3 盈利能力分析', 'level': 3})
        elements.extend(self._analyze_profitability(ic))

        elements.append({'type': 'heading', 'content': '5.4.4 现金流分析', 'level': 3})
        elements.extend(self._analyze_cashflow(cf))

        return elements

    def _analyze_asset_structure(self, bs: pd.DataFrame) -> List[Dict[str, Any]]:
        elements = []
        if bs.empty:
            elements.append({'type': 'paragraph', 'content': '暂无资产数据。'})
            return elements

        periods = self._get_annual_periods(bs, 2)
        if len(periods) < 1:
            return elements

        latest = self._get_row_for_period(bs, periods[0])
        total_assets = self._val(latest, 'total_assets')

        if total_assets and total_assets > 0:
            cur_assets = self._val(latest, 'total_cur_assets') or 0
            cur_pct = cur_assets / total_assets * 100
            inventories = self._val(latest, 'inventories') or 0
            ar = self._val(latest, 'accounts_receiv') or 0
            money = self._val(latest, 'money_cap') or 0

            text = (
                f"截至最新报告期，公司总资产{total_assets / 10000:,.2f}万元，"
                f"其中流动资产{cur_assets / 10000:,.2f}万元（占比{cur_pct:.1f}%）。"
                f"主要资产科目：货币资金{money / 10000:,.2f}万元、"
                f"应收账款{ar / 10000:,.2f}万元、存货{inventories / 10000:,.2f}万元。"
            )
            elements.append({'type': 'paragraph', 'content': text})

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
            st_borr = self._val(latest, 'st_borr') or 0
            lt_borr = self._val(latest, 'lt_borr') or 0
            interest_debt = st_borr + lt_borr

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

        periods = self._get_annual_periods(ic, 5)
        if not periods:
            return elements

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

        periods = self._get_annual_periods(cf, 5)
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

        dim_table = []
        for dim_key, dim_name in dim_name_map.items():
            dim_data = scoring_result.get('dimensions', {}).get(dim_key, {})
            dim_table.append({
                '维度': dim_name,
                '权重': f"{dim_data.get('weight', 0) * 100:.0f}%",
                '得分': f"{dim_data.get('score', 0):.2f}",
                '加权得分': f"{dim_data.get('weighted_score', 0):.2f}",
            })

        elements.append({'type': 'table', 'headers': ['维度', '权重', '得分', '加权得分'], 'data': dim_table})

        total = scoring_result.get('total_score', 0)
        grade = scoring_result.get('grade', '暂无数据')
        elements.append({'type': 'paragraph', 'content': f"综合财务评分{total:.2f}分，评级为\"{grade}\"。"})

        for dim_key, dim_name in dim_name_map.items():
            indicators = scoring_result.get('indicators', {}).get(dim_key, [])
            if not indicators:
                continue

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

            elements.append({'type': 'paragraph', 'content': f"{dim_name}："})
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

        comparison_data = []
        for field in field_names:
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
            better_count = (peer_values < company_value).sum()
        else:
            better_count = (peer_values > company_value).sum()

        rank = int(better_count) + 1
        rank = min(rank, len(peer_values))

        return f"{rank}/{len(peer_values)}"

    # ============================================================
    # 5.7 分红情况分析
    # ============================================================
    def _generate_dividend_analysis(self) -> List[Dict[str, Any]]:
        elements = []
        dividend = self.data.get('dividend', pd.DataFrame())

        if dividend.empty:
            elements.append({'type': 'paragraph', 'content': '暂无分红数据。'})
            return elements

        if 'end_date' not in dividend.columns:
            elements.append({'type': 'paragraph', 'content': '分红数据格式异常。'})
            return elements

        dividend = dividend.copy()
        dividend['year'] = dividend['end_date'].astype(str).str[:4]
        years = sorted(dividend['year'].unique(), reverse=True)[:5]

        table_data = []
        for year in years:
            year_data = dividend[dividend['year'] == year]

            cash_div = None
            for _, row in year_data.iterrows():
                cash_ps = row.get('cash_div', None) or row.get('cash_div_tax', None)
                if cash_ps and pd.notna(cash_ps):
                    cash_div = float(cash_ps)
                    break

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
                elements.append({'type': 'paragraph', 'content': '近年内公司未进行现金分红。'})

        return elements

    # ============================================================
    # 5.8 财务异常汇总
    # ============================================================
    def _generate_anomaly_summary(self) -> List[Dict[str, Any]]:
        """检测两类异常：占比变动异常和行业评级极差"""
        elements = []
        anomalies = []

        # 检查三张报表
        statements = [
            ('资产负债表', 'balance_sheet', self.BS_ASSET_ITEMS + self.BS_LIABILITY_ITEMS + self.BS_EQUITY_ITEMS,
             'total_assets', False),
            ('利润表', 'income_statement',
             [(f, n) for f, n in self.IS_ITEMS if f not in ('total_revenue', 'basic_eps')],
             'total_revenue', False),
            ('现金流量表', 'cashflow',
             [(f, n) for f, n in self.CF_ITEMS if f != 'n_cashflow_act'],
             'n_cashflow_act', True),
        ]

        for stmt_name, stmt_key, items, denom_field, use_abs in statements:
            df = self.data.get('financial_statements', {}).get(stmt_key, pd.DataFrame())
            if df.empty:
                continue

            periods = self._get_annual_periods(df, 5)
            if len(periods) < 2:
                continue

            peer_df = self._get_peer_df(stmt_name)

            for field, name in items:
                # 1. 占比变动异常：最新年 vs 上年，变动超过50个百分点
                latest_val = self._get_value_for_period(df, field, periods[0])
                prev_val = self._get_value_for_period(df, field, periods[1])
                latest_denom = self._get_value_for_period(df, denom_field, periods[0])
                prev_denom = self._get_value_for_period(df, denom_field, periods[1])

                if (latest_val is not None and prev_val is not None
                        and latest_denom is not None and prev_denom is not None):
                    d1 = abs(latest_denom) if use_abs else latest_denom
                    d2 = abs(prev_denom) if use_abs else prev_denom
                    if d1 != 0 and d2 != 0:
                        pct_latest = latest_val / d1 * 100
                        pct_prev = prev_val / d2 * 100
                        change = abs(pct_latest - pct_prev)
                        if change > 50:
                            anomalies.append({
                                '报表': stmt_name,
                                '科目': name,
                                '异常类型': '占比变动异常',
                                '详情': f"占比从{pct_prev:.1f}%变为{pct_latest:.1f}%（变动{change:.1f}个百分点）"
                            })

                # 2. 行业评级极差：百分位 < 20%
                if peer_df is not None and not peer_df.empty:
                    pct = self._calc_item_percentile(df, peer_df, field, denom_field,
                                                      periods[0], use_abs)
                    if pct is not None and pct < 20:
                        anomalies.append({
                            '报表': stmt_name,
                            '科目': name,
                            '异常类型': '行业评级极差',
                            '详情': f"行业百分位{pct:.0f}%（评级：{self._percentile_to_rating(pct)}）"
                        })

        if anomalies:
            elements.append({
                'type': 'table',
                'headers': ['报表', '科目', '异常类型', '详情'],
                'data': anomalies
            })
        else:
            elements.append({
                'type': 'paragraph',
                'content': '经检测，公司各报表科目占比变动正常，行业评级无明显异常。'
            })

        return elements

    # ============================================================
    # 工具方法
    # ============================================================
    def _get_annual_periods(self, df: pd.DataFrame, n: int = 5) -> list:
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
