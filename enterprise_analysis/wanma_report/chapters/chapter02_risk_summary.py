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

        # 2.1 财务状况评级
        elements.append({
            'type': 'heading',
            'content': '2.1 财务状况评级',
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

        # 计算估值风险评级
        val_risk_rating, val_risk_level = self._assess_valuation_risk(valuation_data)

        # 计算退出风险评级
        exit_risk_rating, exit_risk_level = self._assess_exit_risk(scoring_result, valuation_data)

        # 风险评级表格
        risk_rating = [
            {'维度': '综合财务评分', '评级': f"{total_score:.2f}分 ({grade})"},
            {'维度': '估值风险', '评级': val_risk_rating},
            {'维度': '退出风险', '评级': exit_risk_rating},
        ]

        elements.append({
            'type': 'table',
            'headers': ['维度', '评级'],
            'data': risk_rating
        })

        return elements

    def _assess_valuation_risk(self, valuation_data: Dict) -> tuple:
        """评估估值风险

        Args:
            valuation_data: 估值数据

        Returns:
            (评级描述, 风险等级)
        """
        pe_ratio = valuation_data.get('pe_ratio')
        pb_ratio = valuation_data.get('pb_ratio')
        market_cap = valuation_data.get('market_cap')

        if not pe_ratio and not pb_ratio:
            return '暂无估值数据', '待评估'

        risk_score = 0

        # PE评估
        if pe_ratio:
            if pe_ratio < 0:
                risk_score += 3  # 亏损
            elif pe_ratio < 20:
                risk_score += 1  # 低估
            elif pe_ratio < 40:
                risk_score += 2  # 合理
            elif pe_ratio < 60:
                risk_score += 3  # 偏高
            else:
                risk_score += 4  # 高估

        # PB评估
        if pb_ratio:
            if pb_ratio < 1:
                risk_score += 1
            elif pb_ratio < 3:
                risk_score += 2
            elif pb_ratio < 5:
                risk_score += 3
            else:
                risk_score += 4

        # 综合判定
        avg_score = risk_score / 2 if (pe_ratio and pb_ratio) else risk_score

        if avg_score <= 1.5:
            return f"PE:{pe_ratio:.1f}" if pe_ratio else f"PB:{pb_ratio:.1f}", '低风险'
        elif avg_score <= 2.5:
            desc_parts = []
            if pe_ratio:
                desc_parts.append(f"PE:{pe_ratio:.1f}")
            if pb_ratio:
                desc_parts.append(f"PB:{pb_ratio:.1f}")
            return '/'.join(desc_parts), '中等风险'
        elif avg_score <= 3.5:
            desc_parts = []
            if pe_ratio:
                desc_parts.append(f"PE:{pe_ratio:.1f}")
            if pb_ratio:
                desc_parts.append(f"PB:{pb_ratio:.1f}")
            return '/'.join(desc_parts), '较高风险'
        else:
            desc_parts = []
            if pe_ratio:
                desc_parts.append(f"PE:{pe_ratio:.1f}")
            if pb_ratio:
                desc_parts.append(f"PB:{pb_ratio:.1f}")
            return '/'.join(desc_parts), '高风险'

    def _assess_exit_risk(self, scoring_result: Dict, valuation_data: Dict) -> tuple:
        """评估退出风险

        Args:
            scoring_result: 评分结果
            valuation_data: 估值数据

        Returns:
            (评级描述, 风险等级)
        """
        risk_factors = 0
        total_factors = 0

        # 因子1：财务评分（越低退出风险越高）
        total_score = scoring_result.get('total_score', 0)
        if total_score >= 80:
            risk_factors += 1
        elif total_score >= 70:
            risk_factors += 2
        elif total_score >= 60:
            risk_factors += 3
        else:
            risk_factors += 4
        total_factors += 1

        # 因子2：估值水平（越高退出风险越高）
        pe_ratio = valuation_data.get('pe_ratio')
        if pe_ratio:
            if pe_ratio < 20:
                risk_factors += 1
            elif pe_ratio < 40:
                risk_factors += 2
            else:
                risk_factors += 3
            total_factors += 1

        # 因子3：流动性（换手率）
        turnover = valuation_data.get('turnover_rate')
        if turnover:
            if turnover > 3:
                risk_factors += 1
            elif turnover > 1:
                risk_factors += 2
            else:
                risk_factors += 3
            total_factors += 1

        if total_factors == 0:
            return '数据不足', '待评估'

        avg_risk = risk_factors / total_factors
        if avg_risk <= 1.5:
            return '退出渠道畅通', '低风险'
        elif avg_risk <= 2.5:
            return '退出条件一般', '中等风险'
        else:
            return '退出存在不确定性', '较高风险'

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
                })

        if risk_points:
            elements.append({
                'type': 'table',
                'headers': ['风险类型', '风险描述'],
                'data': risk_points
            })
        else:
            elements.append({
                'type': 'paragraph',
                'content': '公司各项财务指标表现良好，未发现明显风险点。'
            })

        return elements

    def _generate_risk_suggestions(self) -> List[Dict[str, Any]]:
        """生成风险应对建议（基于实际评分数据动态生成）"""
        elements = []

        scoring_result = self.data.get('scoring_result', {})
        dimensions = scoring_result.get('dimensions', {})

        suggestions = []

        # 财务风险建议
        total_score = scoring_result.get('total_score', 0)
        if total_score < 70:
            weak_dims = []
            for dim, data in dimensions.items():
                if data.get('score', 0) < 70:
                    dim_name = self._get_dimension_name(dim)
                    weak_dims.append(dim_name)
            if weak_dims:
                suggestions.append({
                    '风险类型': '财务风险',
                    '应对建议': f"关注{'+'.join(weak_dims)}方面，定期跟踪核心财务指标变化"
                })
            else:
                suggestions.append({
                    '风险类型': '财务风险',
                    '应对建议': '持续关注公司财务状况，定期进行财务分析'
                })
        else:
            suggestions.append({
                '风险类型': '财务风险',
                '应对建议': '公司财务状况整体健康，建议定期跟踪关键指标'
            })

        # 估值风险建议
        valuation_data = self.data.get('valuation_data', {})
        pe_ratio = valuation_data.get('pe_ratio')
        if pe_ratio and pe_ratio > 40:
            suggestions.append({
                '风险类型': '估值风险',
                '应对建议': f"当前PE({pe_ratio:.1f})偏高，建议审慎评估定增价格，关注市场估值回调风险"
            })
        else:
            suggestions.append({
                '风险类型': '估值风险',
                '应对建议': '合理评估定增价格，关注行业估值水平变化'
            })

        # 退出风险建议
        suggestions.append({
            '风险类型': '退出风险',
            '应对建议': '制定多元化退出策略（二级市场减持、大宗交易、协议转让），分散减持时点降低冲击成本'
        })

        # 运营风险建议
        operations_score = dimensions.get('operations', {}).get('score', 0)
        if operations_score < 70:
            suggestions.append({
                '风险类型': '运营风险',
                '应对建议': '重点关注公司存货周转和应收账款回收情况，防范运营效率下降风险'
            })

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
