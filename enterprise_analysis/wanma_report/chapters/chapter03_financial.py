#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第三章：财务风险分析
包含基于EFAES评分系统的财务评分
"""

from typing import Dict, Any, List
import pandas as pd


class Chapter03Financial:
    """财务风险分析章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any], scoring_engine):
        """初始化

        Args:
            config: 配置信息
            data: 报告数据
            scoring_engine: 财务评分引擎
        """
        self.config = config
        self.data = data
        self.scoring_engine = scoring_engine
        self.project_info = config.get('project', {})

    def generate(self) -> List[Dict[str, Any]]:
        """生成章节内容

        Returns:
            章节元素列表
        """
        elements = []

        # 章节标题
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
            'content': '3.2 运营能力分析',
            'level': 2
        })

        elements.extend(self._generate_operations_analysis())

        # 3.3 盈利能力分析
        elements.append({
            'type': 'heading',
            'content': '3.3 盈利能力分析',
            'level': 2
        })

        elements.extend(self._generate_profitability_analysis())

        # 3.4 偿债能力分析
        elements.append({
            'type': 'heading',
            'content': '3.4 偿债能力分析',
            'level': 2
        })

        elements.extend(self._generate_solvency_analysis())

        # 3.5 成长能力分析
        elements.append({
            'type': 'heading',
            'content': '3.5 成长能力分析',
            'level': 2
        })

        elements.extend(self._generate_growth_analysis())

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

        # 综合评分表
        score_summary = [
            {'指标': '综合评分', '数值': f"{scoring_result.get('total_score', 0):.2f}分", '说明': scoring_result.get('grade', '')},
            {'指标': '运营能力', '数值': f"{scoring_result.get('dimensions', {}).get('operations', {}).get('score', 0):.2f}分", '说明': ''},
            {'指标': '盈利能力', '数值': f"{scoring_result.get('dimensions', {}).get('profitability', {}).get('score', 0):.2f}分", '说明': ''},
            {'指标': '偿债能力', '数值': f"{scoring_result.get('dimensions', {}).get('solvency', {}).get('score', 0):.2f}分", '说明': ''},
            {'指标': '成长能力', '数值': f"{scoring_result.get('dimensions', {}).get('growth', {}).get('score', 0):.2f}分", '说明': ''},
        ]

        elements.append({
            'type': 'table',
            'headers': ['指标', '数值', '说明'],
            'data': score_summary
        })

        # 评分等级说明
        grade_desc = f"""
财务评分采用EFAES评分系统，综合考虑运营能力、盈利能力、偿债能力和成长能力四个维度。

评分标准：
- 优秀(90-100分)：各项财务指标表现优异，处于行业领先水平
- 良好(80-89分)：财务指标表现良好，具有较强的竞争力
- 中等(70-79分)：财务指标表现一般，处于行业中游水平
- 较差(60-69分)：部分财务指标较弱，需要关注改进
- 极差(0-59分)：多项财务指标表现较差，存在较大风险

公司当前综合评分为{scoring_result.get('total_score', 0):.2f}分，评级为"{scoring_result.get('grade', '暂无数据')}"。
"""
        elements.append({
            'type': 'paragraph',
            'content': grade_desc.strip()
        })

        return elements

    def _generate_operations_analysis(self) -> List[Dict[str, Any]]:
        """生成运营能力分析"""
        elements = []

        scoring_result = self.data.get('scoring_result', {})
        indicators = scoring_result.get('indicators', {}).get('operations', [])

        if indicators:
            table_data = []
            for ind in indicators:
                table_data.append({
                    '指标': ind.get('name', ''),
                    '数值': ind.get('value'),
                    '评分': ind.get('score'),
                    '状态': ind.get('status', '')
                })

            elements.append({
                'type': 'table',
                'headers': ['指标', '数值', '评分', '状态'],
                'data': table_data
            })

        return elements

    def _generate_profitability_analysis(self) -> List[Dict[str, Any]]:
        """生成盈利能力分析"""
        elements = []

        scoring_result = self.data.get('scoring_result', {})
        indicators = scoring_result.get('indicators', {}).get('profitability', [])

        if indicators:
            table_data = []
            for ind in indicators:
                table_data.append({
                    '指标': ind.get('name', ''),
                    '数值': ind.get('value'),
                    '评分': ind.get('score'),
                    '状态': ind.get('status', '')
                })

            elements.append({
                'type': 'table',
                'headers': ['指标', '数值', '评分', '状态'],
                'data': table_data
            })

        return elements

    def _generate_solvency_analysis(self) -> List[Dict[str, Any]]:
        """生成偿债能力分析"""
        elements = []

        scoring_result = self.data.get('scoring_result', {})
        indicators = scoring_result.get('indicators', {}).get('solvency', [])

        if indicators:
            table_data = []
            for ind in indicators:
                table_data.append({
                    '指标': ind.get('name', ''),
                    '数值': ind.get('value'),
                    '评分': ind.get('score'),
                    '状态': ind.get('status', '')
                })

            elements.append({
                'type': 'table',
                'headers': ['指标', '数值', '评分', '状态'],
                'data': table_data
            })

        return elements

    def _generate_growth_analysis(self) -> List[Dict[str, Any]]:
        """生成成长能力分析"""
        elements = []

        scoring_result = self.data.get('scoring_result', {})
        indicators = scoring_result.get('indicators', {}).get('growth', [])

        if indicators:
            table_data = []
            for ind in indicators:
                table_data.append({
                    '指标': ind.get('name', ''),
                    '数值': ind.get('value'),
                    '评分': ind.get('score'),
                    '状态': ind.get('status', '')
                })

            elements.append({
                'type': 'table',
                'headers': ['指标', '数值', '评分', '状态'],
                'data': table_data
            })

        return elements

    def _generate_industry_comparison(self) -> List[Dict[str, Any]]:
        """生成行业对比分析"""
        elements = []

        industry_data = self.data.get('industry_data', pd.DataFrame())

        if not industry_data.empty:
            # 获取主要指标的行业对比
            key_fields = ['roe', 'gross_margin', 'net_margin', 'current_ratio', 'debt_to_assets']

            comparison_data = []
            for field in key_fields:
                if field in industry_data.columns:
                    industry_median = industry_data[field].median()
                    company_value = self._get_company_value(field)

                    comparison_data.append({
                        '指标': field,
                        '公司值': company_value,
                        '行业中位数': round(industry_median, 4) if industry_median is not None else None,
                        '行业排名': '[待计算]'
                    })

            if comparison_data:
                elements.append({
                    'type': 'table',
                    'headers': ['指标', '公司值', '行业中位数', '行业排名'],
                    'data': comparison_data
                })

        return elements

    def _get_company_value(self, field: str):
        """获取公司指标值"""
        financial_indicators = self.data.get('financial_indicators', pd.DataFrame())
        if financial_indicators.empty or field not in financial_indicators.columns:
            return None

        # 获取最新一期数据
        df_sorted = financial_indicators.sort_values('end_date', ascending=False)
        if len(df_sorted) > 0:
            value = df_sorted.iloc[0].get(field)
            return round(value, 4) if value is not None else None

        return None
