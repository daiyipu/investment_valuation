#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第六章：招引落地
"""

from typing import Dict, Any, List


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

        background_text = f"""
{company_info.get('name', '该公司')}拟通过定向增发方式募集资金，用于扩大产能和补充流动资金。
本次定增项目是公司发展的重要战略举措，有助于提升公司核心竞争力，扩大市场份额。
"""
        elements.append({
            'type': 'paragraph',
            'content': background_text.strip()
        })

        return elements

    def _generate_cooperation_mode(self) -> List[Dict[str, Any]]:
        """生成合作模式"""
        elements = []

        cooperation_modes = [
            {'模式': '参与定增', '说明': '作为战略投资者参与公司定向增发', '权益': '获得定增股份，享有锁定期满后的减持权利'},
            {'模式': '业绩承诺', '说明': '可要求原股东或公司提供业绩承诺', '权益': '保障投资回报，降低业绩风险'},
            {'模式': '股份回购', '说明': '可约定条件下的股份回购条款', '权益': '提供退出保障'},
        ]

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
            {'条件类型': '合规条件', '具体要求': '完成监管部门要求的审批程序'},
            {'条件类型': '协议条件', '具体要求': '签署股份认购协议及相关配套协议'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['条件类型', '具体要求'],
            'data': conditions
        })

        return elements

    def _generate_risk_assurance(self) -> List[Dict[str, Any]]:
        """生成风险保障"""
        elements = []

        assurances = [
            {'保障措施': '股份锁定', '内容': '定增股份按监管要求进行锁定'},
            {'保障措施': '信息披露', '内容': '要求公司及时、准确披露相关信息'},
            {'保障措施': '业绩监控', '内容': '定期跟踪公司经营状况和业绩表现'},
            {'保障措施': '退出预案', '内容': '制定多元化退出策略和预案'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['保障措施', '内容'],
            'data': assurances
        })

        return elements
