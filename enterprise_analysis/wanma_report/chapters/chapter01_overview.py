#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第一章：项目概述
"""

from typing import Dict, Any, List


class Chapter01Overview:
    """项目概述章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any]):
        self.config = config
        self.data = data
        self.project_info = config.get('project', {})

    def generate(self) -> List[Dict[str, Any]]:
        elements = []

        elements.append({
            'type': 'heading',
            'content': '第一章 项目概述',
            'level': 1
        })

        # 1.1 项目基本信息
        elements.append({
            'type': 'heading',
            'content': '1.1 项目基本信息',
            'level': 2
        })
        elements.extend(self._generate_project_basic_info())

        # 1.2 定增项目概况
        elements.append({
            'type': 'heading',
            'content': '1.2 定增项目概况',
            'level': 2
        })
        elements.extend(self._generate_dingzeng_info())

        return elements

    def _generate_project_basic_info(self) -> List[Dict[str, Any]]:
        elements = []
        project_info = self.project_info
        company_info = self.data.get('company_info', {})

        basic_info = [
            {'项目': '项目名称', '内容': f"{project_info.get('name', '')}定增项目"},
            {'项目': '股票代码', '内容': project_info.get('stock_code', '')},
            {'项目': '公司名称', '内容': company_info.get('name', '')},
            {'项目': '所属行业', '内容': company_info.get('industry', '')},
            {'项目': '上市板块', '内容': company_info.get('market', '')},
            {'项目': '上市日期', '内容': company_info.get('list_date', '')},
            {'项目': '报告日期', '内容': project_info.get('report_date', '')},
            {'项目': '报告版本', '内容': project_info.get('version', '')},
        ]

        elements.append({
            'type': 'table',
            'headers': ['项目', '内容'],
            'data': basic_info
        })

        return elements

    def _generate_dingzeng_info(self) -> List[Dict[str, Any]]:
        elements = []
        dingzeng_config = self.config.get('dingzeng', {})

        dingzeng_info = [
            {'项目': '定增目的', '内容': dingzeng_config.get('purpose', '扩大产能、补充流动资金')},
            {'项目': '拟募集资金', '内容': dingzeng_config.get('fundraising_amount', '[待填入]')},
            {'项目': '发行价格', '内容': dingzeng_config.get('issue_price', '[待填入]')},
            {'项目': '发行数量', '内容': dingzeng_config.get('issue_shares', '[待填入]')},
        ]

        elements.append({
            'type': 'table',
            'headers': ['项目', '内容'],
            'data': dingzeng_info
        })

        return elements
