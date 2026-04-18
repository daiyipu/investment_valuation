#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第六章：招引落地
"""

from typing import Dict, Any, List
import pandas as pd


class Chapter06Investment:
    """招引落地章节生成器"""

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
            'content': '第六章 招引落地',
            'level': 1
        })

        # 6.1 项目背景
        elements.append({
            'type': 'heading',
            'content': '6.1 项目背景',
            'level': 2
        })

        elements.extend(self._generate_project_background())

        # 6.2 合作模式
        elements.append({
            'type': 'heading',
            'content': '6.2 合作模式',
            'level': 2
        })

        elements.extend(self._generate_cooperation_mode())

        # 6.3 落地条件
        elements.append({
            'type': 'heading',
            'content': '6.3 落地条件',
            'level': 2
        })

        elements.extend(self._generate_landing_conditions())

        # 6.4 风险保障
        elements.append({
            'type': 'heading',
            'content': '6.4 风险保障',
            'level': 2
        })

        elements.extend(self._generate_risk_assurance())

        return elements

    def _generate_project_background(self) -> List[Dict[str, Any]]:
        """生成项目背景"""
        elements = []

        company_info = self.data.get('company_info', {})
        scoring_result = self.data.get('scoring_result', {})
        valuation_data = self.data.get('valuation_data', {})
        dingzeng_config = self.config.get('dingzeng', {})

        company_name = company_info.get('name', '该公司')
        industry = company_info.get('industry', '')
        total_score = scoring_result.get('total_score', 0)
        grade = scoring_result.get('grade', '')
        market_cap = valuation_data.get('market_cap')

        # 构建项目背景描述
        bg_parts = [
            f'{company_name}（股票代码：{self.project_info.get("stock_code", "")}）',
            f'属于{industry}行业，' if industry else '',
        ]

        background_text = ''.join(bg_parts)
        background_text += f'拟通过定向增发方式募集资金，{dingzeng_config.get("purpose", "用于公司业务发展及补充流动资金")}。'

        if total_score > 0:
            background_text += f'\n\n公司当前综合财务评分为{total_score:.2f}分（{grade}），'

            if total_score >= 80:
                background_text += '财务状况良好，具有较强的投资价值。'
            elif total_score >= 70:
                background_text += '财务状况处于行业中游水平，具备一定投资价值。'
            else:
                background_text += '财务状况存在一定风险，需审慎评估投资价值。'

        if market_cap:
            background_text += f'\n\n公司当前总市值约{market_cap / 10000:.2f}亿元。'

        elements.append({
            'type': 'paragraph',
            'content': background_text.strip()
        })

        return elements

    def _generate_cooperation_mode(self) -> List[Dict[str, Any]]:
        """生成合作模式"""
        elements = []

        scoring_result = self.data.get('scoring_result', {})
        total_score = scoring_result.get('total_score', 0)

        cooperation_modes = [
            {'模式': '参与定增', '说明': '作为战略投资者参与公司定向增发', '权益': '获得定增股份，享有锁定期满后的减持权利'},
            {'模式': '业绩承诺', '说明': '可要求原股东或公司提供业绩承诺', '权益': '保障投资回报，降低业绩不达预期风险'},
            {'模式': '股份回购', '说明': '可约定特定条件下的股份回购条款', '权益': '提供退出保障，降低投资损失风险'},
        ]

        # 根据财务评分动态调整建议
        if total_score < 70:
            cooperation_modes.append({
                '模式': '对赌条款',
                '说明': '设定业绩对赌条件，未达标时原股东进行补偿',
                '权益': '强化投资保障，降低财务风险'
            })

        elements.append({
            'type': 'table',
            'headers': ['模式', '说明', '权益'],
            'data': cooperation_modes
        })

        return elements

    def _generate_landing_conditions(self) -> List[Dict[str, Any]]:
        """生成落地条件"""
        elements = []

        conditions = [
            {'条件类型': '资金条件', '具体要求': '按时足额缴纳认购款项'},
            {'条件类型': '资质条件', '具体要求': '符合上市公司定向增发的投资者适当性要求'},
            {'条件类型': '合规条件', '具体要求': '完成监管部门要求的审批程序（证监会注册等）'},
            {'条件类型': '协议条件', '具体要求': '签署股份认购协议及相关配套协议'},
        ]

        # 根据估值数据添加估值条件
        valuation_data = self.data.get('valuation_data', {})
        pe_ratio = valuation_data.get('pe_ratio')
        if pe_ratio:
            conditions.append({
                '条件类型': '估值条件',
                '具体要求': f'定增发行价格应合理（当前二级市场PE约{pe_ratio:.1f}倍），折价率需在合理范围内'
            })

        elements.append({
            'type': 'table',
            'headers': ['条件类型', '具体要求'],
            'data': conditions
        })

        return elements

    def _generate_risk_assurance(self) -> List[Dict[str, Any]]:
        """生成风险保障"""
        elements = []

        scoring_result = self.data.get('scoring_result', {})
        dimensions = scoring_result.get('dimensions', {})
        total_score = scoring_result.get('total_score', 0)

        assurances = [
            {'保障措施': '股份锁定', '内容': '定增股份按监管要求进行锁定（通常6个月），锁定期内不可减持'},
            {'保障措施': '信息披露', '内容': '要求公司按监管要求及时、准确、完整披露定期报告和临时公告'},
            {'保障措施': '业绩监控', '内容': '定期跟踪公司经营状况和业绩表现，关注核心财务指标变化'},
            {'保障措施': '退出预案', '内容': '制定多元化退出策略（二级市场+大宗交易+协议转让），分散减持时点'},
        ]

        # 根据风险维度添加针对性保障措施
        dim_names = {
            'operations': '运营效率',
            'profitability': '盈利能力',
            'solvency': '偿债能力',
            'growth': '成长能力'
        }

        for dim, data in dimensions.items():
            score = data.get('score', 0)
            if score < 60:
                dim_name = dim_names.get(dim, dim)
                assurances.append({
                    '保障措施': f'{dim_name}预警',
                    'content': f'{dim_name}评分仅{score:.1f}分，建议设置专项监控指标，出现恶化趋势时及时预警'
                })

        elements.append({
            'type': 'table',
            'headers': ['保障措施', '内容'],
            'data': assurances
        })

        # 综合风险提示
        if total_score < 70:
            elements.append({
                'type': 'paragraph',
                'content': f'综合提示：公司财务评分{total_score:.1f}分，处于较低水平，建议在投资协议中强化保障条款，包括但不限于业绩承诺、股份回购、对赌安排等。'
            })

        return elements
