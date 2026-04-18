#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第三章：财务风险分析
基于EFAES评分系统，27项指标，四维度评分
"""

from typing import Dict, Any, List
import pandas as pd


class Chapter03Financial:
    """财务风险分析章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any], scoring_engine):
        self.config = config
        self.data = data
        self.scoring_engine = scoring_engine
        self.project_info = config.get('project', {})

    def generate(self) -> List[Dict[str, Any]]:
        elements = []

        elements.append({
            'type': 'heading',
            'content': '第三章 财务风险分析',
            'level': 1
        })

        # 3.1 综合财务评分
        elements.append({
            'type': 'heading',
            'content': '3.1 综合财务评分',
            'level': 2
        })
        elements.extend(self._generate_comprehensive_score())

        # 3.2 运营能力分析
        elements.append({
            'type': 'heading',
            'content': '3.2 运营能力分析（权重15%）',
            'level': 2
        })
        elements.extend(self._generate_dimension_analysis('operations'))

        # 3.3 盈利能力分析
        elements.append({
            'type': 'heading',
            'content': '3.3 盈利能力分析（权重30%）',
            'level': 2
        })
        elements.extend(self._generate_dimension_analysis('profitability'))

        # 3.4 偿债能力分析
        elements.append({
            'type': 'heading',
            'content': '3.4 偿债能力分析（权重30%）',
            'level': 2
        })
        elements.extend(self._generate_solvency_analysis())

        # 3.5 成长能力分析
        elements.append({
            'type': 'heading',
            'content': '3.5 成长能力分析（权重25%）',
            'level': 2
        })
        elements.extend(self._generate_dimension_analysis('growth'))

        # 3.6 行业对比分析
        elements.append({
            'type': 'heading',
            'content': '3.6 行业对比分析',
            'level': 2
        })
        elements.extend(self._generate_industry_comparison())

        return elements

    def _generate_comprehensive_score(self) -> List[Dict[str, Any]]:
        """生成综合财务评分"""
        elements = []
        scoring_result = self.data.get('scoring_result', {})

        # 维度汇总表
        dim_table = []
        dim_name_map = {
            'operations': '运营能力',
            'profitability': '盈利能力',
            'solvency': '偿债能力',
            'growth': '成长能力',
        }

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

        # 综合结论
        total = scoring_result.get('total_score', 0)
        grade = scoring_result.get('grade', '暂无数据')

        # 找出最强和最弱维度
        dim_scores = {
            dim_name_map[k]: v.get('score', 0)
            for k, v in scoring_result.get('dimensions', {}).items()
        }
        if dim_scores:
            strongest = max(dim_scores, key=dim_scores.get)
            weakest = min(dim_scores, key=dim_scores.get)
        else:
            strongest = weakest = '-'

        grade_desc = f"""
财务评分采用EFAES评分系统，涵盖27项财务指标，综合评估运营能力（15%）、盈利能力（30%）、偿债能力（30%）和成长能力（25%）四个维度。

评分方法：将各指标值与行业分位数阈值（p10/p30/p50/p70/p90）对比，分为优秀（0.9）、良好（0.8）、中等（0.7）、较差（0.6）、极差（0.5）五个等级，加权汇总得出综合评分。

公司当前综合评分为{total:.2f}分，评级为"{grade}"。
其中{strongest}表现最强（{dim_scores.get(strongest, 0):.2f}分），{weakest}表现最弱（{dim_scores.get(weakest, 0):.2f}分）。
"""
        elements.append({
            'type': 'paragraph',
            'content': grade_desc.strip()
        })

        return elements

    def _generate_dimension_analysis(self, dimension: str) -> List[Dict[str, Any]]:
        """生成通用维度分析"""
        elements = []
        scoring_result = self.data.get('scoring_result', {})
        indicators = scoring_result.get('indicators', {}).get(dimension, [])

        if indicators:
            table_data = []
            for ind in indicators:
                value_str = f"{ind.get('value'):.4f}" if ind.get('value') is not None else '-'
                score_str = f"{ind.get('score'):.1f}" if ind.get('score') is not None else '-'
                ratio_str = ind.get('quantile_level', '-')

                table_data.append({
                    '指标': ind.get('name', ''),
                    '数值': value_str,
                    '等级': ratio_str,
                    '得分': score_str,
                    '状态': ind.get('status', ''),
                })

            elements.append({
                'type': 'table',
                'headers': ['指标', '数值', '等级', '得分', '状态'],
                'data': table_data
            })

            # 维度分析小结
            valid_scores = [i['score'] for i in indicators if i.get('score') is not None]
            if valid_scores:
                avg_score = sum(valid_scores) / len(valid_scores)
                weak_items = [i for i in indicators if i.get('score') is not None and i['score'] < 70]
                strong_items = [i for i in indicators if i.get('score') is not None and i['score'] >= 90]

                summary = f"本维度平均得分{avg_score:.1f}分。"
                if strong_items:
                    names = '、'.join(i['name'] for i in strong_items)
                    summary += f"其中{names}表现优秀。"
                if weak_items:
                    names = '、'.join(i['name'] for i in weak_items)
                    summary += f"需要关注{names}的改善。"

                elements.append({
                    'type': 'paragraph',
                    'content': summary
                })

        return elements

    def _generate_solvency_analysis(self) -> List[Dict[str, Any]]:
        """生成偿债能力分析（含短期/长期子维度）"""
        elements = []
        scoring_result = self.data.get('scoring_result', {})
        sub_dims = scoring_result.get('sub_dimensions', {}).get('solvency', {})

        # 3.4.1 短期偿债
        elements.append({
            'type': 'heading',
            'content': '3.4.1 短期偿债能力分析',
            'level': 3
        })

        short_term = sub_dims.get('short_term', {})
        short_indicators = short_term.get('indicators', [])
        if short_indicators:
            table_data = self._format_indicator_table(short_indicators)
            elements.append({
                'type': 'table',
                'headers': ['指标', '数值', '等级', '得分', '状态'],
                'data': table_data
            })
            self._add_dimension_comment(elements, short_indicators, '短期偿债')

        # 3.4.2 长期偿债
        elements.append({
            'type': 'heading',
            'content': '3.4.2 长期偿债能力分析',
            'level': 3
        })

        long_term = sub_dims.get('long_term', {})
        long_indicators = long_term.get('indicators', [])
        if long_indicators:
            table_data = self._format_indicator_table(long_indicators)
            elements.append({
                'type': 'table',
                'headers': ['指标', '数值', '等级', '得分', '状态'],
                'data': table_data
            })
            self._add_dimension_comment(elements, long_indicators, '长期偿债')

        return elements

    def _format_indicator_table(self, indicators: list) -> list:
        """格式化指标表格"""
        table_data = []
        for ind in indicators:
            value_str = f"{ind.get('value'):.4f}" if ind.get('value') is not None else '-'
            score_str = f"{ind.get('score'):.1f}" if ind.get('score') is not None else '-'
            table_data.append({
                '指标': ind.get('name', ''),
                '数值': value_str,
                '等级': ind.get('quantile_level', '-'),
                '得分': score_str,
                '状态': ind.get('status', ''),
            })
        return table_data

    def _add_dimension_comment(self, elements: list, indicators: list, dim_name: str) -> None:
        """添加维度分析小结"""
        valid_scores = [i['score'] for i in indicators if i.get('score') is not None]
        if not valid_scores:
            return

        avg_score = sum(valid_scores) / len(valid_scores)
        weak_items = [i for i in indicators if i.get('score') is not None and i['score'] < 70]
        strong_items = [i for i in indicators if i.get('score') is not None and i['score'] >= 90]

        summary = f"{dim_name}维度平均得分{avg_score:.1f}分。"
        if strong_items:
            names = '、'.join(i['name'] for i in strong_items)
            summary += f"其中{names}表现优秀。"
        if weak_items:
            names = '、'.join(i['name'] for i in weak_items)
            summary += f"需要关注{names}的改善。"

        elements.append({
            'type': 'paragraph',
            'content': summary
        })

    def _generate_industry_comparison(self) -> List[Dict[str, Any]]:
        """生成行业对比分析"""
        elements = []

        industry_data = self.data.get('industry_data', pd.DataFrame())
        if industry_data.empty:
            elements.append({
                'type': 'paragraph',
                'content': '暂无行业对比数据。'
            })
            return elements

        # 核心指标对比
        field_names = {
            'roe': '净资产收益率',
            'grossprofit_margin': '毛利率',
            'netprofit_margin': '净利率',
            'current_ratio': '流动比率',
            'debt_to_assets': '资产负债率',
            'assets_turn': '总资产周转率',
            'netprofit_yoy': '净利润增长率',
        }
        # 兼容旧字段名
        alias_to_field = {
            'gross_margin': 'grossprofit_margin',
            'net_margin': 'netprofit_margin',
            'asset_turn': 'assets_turn',
            'net_profit_yoy': 'netprofit_yoy',
        }
        key_fields = ['roe', 'grossprofit_margin', 'netprofit_margin',
                      'current_ratio', 'debt_to_assets', 'assets_turn', 'netprofit_yoy']

        peer_codes = industry_data['ts_code'].unique() if 'ts_code' in industry_data.columns else []
        total_peers = len(peer_codes)

        comparison_data = []
        for field in key_fields:
            col = field
            if col not in industry_data.columns:
                col = alias_to_field.get(field, field)
            if col not in industry_data.columns:
                continue

            display_name = field_names.get(field, field)
            industry_median = industry_data[col].median()
            company_value = self._get_company_value(field)

            ranking = self._calculate_industry_ranking(
                company_value, col, industry_data, total_peers
            )

            comparison_data.append({
                '指标': display_name,
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

    def _calculate_industry_ranking(
        self,
        company_value,
        field: str,
        industry_data: pd.DataFrame,
        total_peers: int,
    ) -> str:
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

        # 1-based排名：正向指标降序（值大为优），反向指标升序（值小为优）
        if is_reverse:
            sorted_values = peer_values.sort_values(ascending=True)
        else:
            sorted_values = peer_values.sort_values(ascending=False)

        rank = 1
        for val in sorted_values:
            if abs(val - company_value) < 1e-10:
                break
            rank += 1
        rank = min(rank, len(sorted_values))

        return f"{rank}/{len(sorted_values)}"

    def _get_company_value(self, field: str):
        financial_indicators = self.data.get('financial_indicators', pd.DataFrame())
        if financial_indicators.empty:
            return None

        col = field
        if col not in financial_indicators.columns:
            # 尝试字段别名
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

        if col not in financial_indicators.columns:
            return None

        df_sorted = financial_indicators.sort_values('end_date', ascending=False)
        if len(df_sorted) > 0:
            value = df_sorted.iloc[0].get(col)
            return round(value, 4) if value is not None else None
        return None
