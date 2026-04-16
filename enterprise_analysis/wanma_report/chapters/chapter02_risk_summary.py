#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第二章：风险概述
"""

from typing import Dict, Any, List


class Chapter02RiskSummary:
    """风险概述章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any]):
        """初始化

        Args:
            config: 配置信息
            data: 报告数据
        """
        self.config = config
        self.data = data
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
            'content': '第二章 风险概述',
            'level': 1
        })

        # 2.1 风险评级
        elements.append({
            'type': 'heading',
            'content': '2.1 风险评级',
            'level': 2
        })

        elements.extend(self._generate_risk_rating())

        # 2.2 主要风险点
        elements.append({
            'type': 'heading',
            'content': '2.2 主要风险点',
            'level': 2
        })

        elements.extend(self._generate_risk_points())

        # 2.3 风险应对建议
        elements.append({
            'type': 'heading',
            'content': '2.3 风险应对建议',
            'level': 2
        })

        elements.extend(self._generate_risk_suggestions())

        return elements

    def _generate_risk_rating(self) -> List[Dict[str, Any]]:
        """生成风险评级"""
        elements = []

        scoring_result = self.data.get('scoring_result', {})
        valuation_data = self.data.get('valuation_data', {})

        total_score = scoring_result.get('total_score', 0)
        grade = scoring_result.get('grade', '暂无数据')

        # 风险评级表格
        risk_rating = [
            {'维度': '综合财务评分', '评级': f"{total_score:.2f}分 ({grade})", '风险等级': self._score_to_risk_level(total_score)},
            {'维度': '估值风险', '评级': '[待评估]', '风险等级': '待评估'},
            {'维度': '退出风险', '评级': '[待评估]', '风险等级': '待评估'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['维度', '评级', '风险等级'],
            'data': risk_rating
        })

        return elements

    def _generate_risk_points(self) -> List[Dict[str, Any]]:
        """生成主要风险点"""
        elements = []

        # 基于财务评分识别风险点
        scoring_result = self.data.get('scoring_result', {})
        dimensions = scoring_result.get('dimensions', {})

        risk_points = []

        # 检查各维度风险
        for dim, data in dimensions.items():
            score = data.get('score', 0)
            if score < 70:
                dim_name = self._get_dimension_name(dim)
                risk_level = self._get_risk_level(score)
                risk_points.append({
                    '风险类型': dim_name,
                    '风险描述': f"{dim_name}评分{score:.2f}分，低于行业平均水平",
                    '风险等级': risk_level
                })

        if risk_points:
            elements.append({
                'type': 'table',
                'headers': ['风险类型', '风险描述', '风险等级'],
                'data': risk_points
            })
        else:
            elements.append({
                'type': 'paragraph',
                'content': '公司各项财务指标表现良好，未发现明显风险点。'
            })

        return elements

    def _generate_risk_suggestions(self) -> List[Dict[str, Any]]:
        """生成风险应对建议"""
        elements = []

        suggestions = [
            {'风险类型': '财务风险', '建议': '持续关注公司财务状况，定期进行财务分析'},
            {'风险类型': '估值风险', '建议': '合理评估定增价格，关注市场估值水平变化'},
            {'风险类型': '退出风险', '建议': '制定多元化退出策略，降低单一退出方式风险'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['风险类型', '应对建议'],
            'data': suggestions
        })

        return elements

    def _score_to_risk_level(self, score: float) -> str:
        """评分转换为风险等级"""
        if score >= 80:
            return '低风险'
        elif score >= 70:
            return '中等风险'
        elif score >= 60:
            return '较高风险'
        else:
            return '高风险'

    def _get_risk_level(self, score: float) -> str:
        """评分转换为风险等级"""
        if score >= 80:
            return '低'
        elif score >= 70:
            return '中'
        elif score >= 60:
            return '较高'
        else:
            return '高'

    def _get_dimension_name(self, dim: str) -> str:
        """获取维度中文名称"""
        names = {
            'operations': '运营能力',
            'profitability': '盈利能力',
            'solvency': '偿债能力',
            'growth': '成长能力'
        }
        return names.get(dim, dim)
