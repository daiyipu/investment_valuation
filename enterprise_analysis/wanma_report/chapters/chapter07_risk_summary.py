#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第七章：风险概述
"""

from typing import Dict, Any, List
from datetime import datetime, timedelta
import pandas as pd


class Chapter07RiskSummary:
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
            'content': '第七章 风险概述',
            'level': 1
        })

        # 2.1 风险评级
        elements.append({
            'type': 'heading',
            'content': '2.1 风险评级',
            'level': 2
        })

        elements.extend(self._generate_risk_rating())

        # 7.2 主要风险点
        elements.append({
            'type': 'heading',
            'content': '7.2 主要风险点',
            'level': 2
        })

        elements.extend(self._generate_risk_points())

        # 7.3 股东交易风险
        elements.append({
            'type': 'heading',
            'content': '7.3 股东交易风险',
            'level': 2
        })
        elements.extend(self._generate_shareholder_trading_risk())

        # 7.4 限售股解禁风险
        elements.append({
            'type': 'heading',
            'content': '7.4 限售股解禁风险',
            'level': 2
        })
        elements.extend(self._generate_share_float_risk())

        # 7.5 风险应对建议
        elements.append({
            'type': 'heading',
            'content': '7.5 风险应对建议',
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

        # 风险评级表格
        risk_rating = [
            {'维度': '综合财务评分', '评级': f"{total_score:.2f}分 ({grade})", '风险等级': self._score_to_risk_level(total_score)},
            {'维度': '估值风险', '评级': val_risk_rating, '风险等级': val_risk_level},
        ]

        elements.append({
            'type': 'table',
            'headers': ['维度', '评级', '风险等级'],
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

    def _generate_shareholder_trading_risk(self) -> List[Dict[str, Any]]:
        """分析大股东交易风险"""
        elements = []
        holdertrade = self.data.get('stk_holdertrade', pd.DataFrame())

        if holdertrade.empty:
            elements.append({'type': 'paragraph', 'content': '暂无大股东交易数据。'})
            return elements

        # 按交易日期排序
        holdertrade = holdertrade.copy()
        if 'ann_date' in holdertrade.columns:
            holdertrade = holdertrade.sort_values('ann_date', ascending=False)

        # 取最近20笔交易
        recent_trades = holdertrade.head(20)

        table_data = []
        for _, row in recent_trades.iterrows():
            ann_date = str(row.get('ann_date', ''))
            if len(ann_date) >= 8:
                ann_date = f"{ann_date[:4]}-{ann_date[4:6]}-{ann_date[6:8]}"

            holder_name = row.get('holder_name', '')
            trade_type = row.get('trade_type', '')
            type_desc = {'B': '买入', 'S': '卖出', '增持': '增持', '减持': '减持'}.get(str(trade_type), str(trade_type))

            vol = row.get('vol', None)
            price = row.get('price', None)
            change_ratio = row.get('change_ratio', None)

            row_data = {
                '公告日期': ann_date,
                '股东名称': str(holder_name),
                '交易类型': type_desc,
                '交易数量(股)': f"{float(vol):,.0f}" if pd.notna(vol) else '-',
                '交易价格': f"{float(price):.2f}" if pd.notna(price) else '-',
                '变动比例': f"{float(change_ratio):.4f}%" if pd.notna(change_ratio) else '-',
            }
            table_data.append(row_data)

        if table_data:
            elements.append({
                'type': 'table',
                'headers': ['公告日期', '股东名称', '交易类型', '交易数量(股)', '交易价格', '变动比例'],
                'data': table_data
            })

            # 买卖统计
            buy_count = sum(1 for r in table_data if r['交易类型'] in ('买入', '增持'))
            sell_count = sum(1 for r in table_data if r['交易类型'] in ('卖出', '减持'))

            text = f"近期大股东交易中，买入/增持{buy_count}笔，卖出/减持{sell_count}笔。"
            if sell_count > buy_count * 2:
                text += "大股东以减持为主，需关注减持对股价的压力。"
            elif buy_count > sell_count * 2:
                text += "大股东以增持为主，显示对公司前景的信心。"
            elements.append({'type': 'paragraph', 'content': text})

        return elements

    def _generate_share_float_risk(self) -> List[Dict[str, Any]]:
        """分析限售股解禁风险"""
        elements = []
        share_float = self.data.get('share_float', pd.DataFrame())

        if share_float.empty:
            elements.append({'type': 'paragraph', 'content': '暂无限售股解禁数据。'})
            return elements

        from datetime import datetime
        today = datetime.now().strftime('%Y%m%d')
        share_float = share_float.copy()

        # 统计未来6个月解禁
        six_months_later = (datetime.now() + timedelta(days=180)).strftime('%Y%m%d')

        if 'float_date' in share_float.columns:
            near_term = share_float[
                (share_float['float_date'].astype(str) >= today) &
                (share_float['float_date'].astype(str) <= six_months_later)
            ]
        else:
            near_term = pd.DataFrame()

        if near_term.empty:
            elements.append({
                'type': 'paragraph',
                'content': '未来6个月内无限售股解禁计划，短期解禁压力较小。'
            })
            return elements

        total_float = near_term['float_share'].sum() if 'float_share' in near_term.columns else 0
        total_ratio = near_term['float_ratio'].sum() if 'float_ratio' in near_term.columns else 0

        if pd.notna(total_ratio) and total_ratio > 0:
            risk_level = '较高风险' if total_ratio > 10 else ('中等风险' if total_ratio > 3 else '低风险')
            text = (
                f"未来6个月内有{len(near_term)}笔限售股解禁，"
                f"合计解禁{total_float:,.0f}股，占总股本{total_ratio:.2f}%。"
                f"解禁风险等级：{risk_level}。"
            )
            elements.append({'type': 'paragraph', 'content': text})
        else:
            elements.append({
                'type': 'paragraph',
                'content': f"未来6个月有{len(near_term)}笔限售股解禁计划，需关注解禁节奏。"
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
