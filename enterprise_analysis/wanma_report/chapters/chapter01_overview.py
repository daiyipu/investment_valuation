#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第一章：项目概况
"""

from typing import Dict, Any, List
import pandas as pd


class Chapter01Overview:
    """项目概况章节生成器"""

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
            'content': '第一章 项目概况',
            'level': 1
        })

        # 1.1 项目基本信息
        elements.append({
            'type': 'heading',
            'content': '1.1 项目基本信息',
            'level': 2
        })

        elements.extend(self._generate_project_basic_info())

        # 1.2 公司基本情况
        elements.append({
            'type': 'heading',
            'content': '1.2 公司基本情况',
            'level': 2
        })

        elements.extend(self._generate_company_basic_info())

        # 1.3 主营业务情况
        elements.append({
            'type': 'heading',
            'content': '1.3 主营业务情况',
            'level': 2
        })

        elements.extend(self._generate_main_business_info())

        # 1.4 定增项目概况
        elements.append({
            'type': 'heading',
            'content': '1.4 定增项目概况',
            'level': 2
        })

        elements.extend(self._generate_dingzeng_info())

        return elements

    def _generate_project_basic_info(self) -> List[Dict[str, Any]]:
        """生成项目基本信息"""
        elements = []

        project_info = self.project_info
        company_info = self.data.get('company_info', {})

        # 基本信息表格
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

    def _generate_company_basic_info(self) -> List[Dict[str, Any]]:
        """生成公司基本情况"""
        elements = []

        company_info = self.data.get('company_info', {})

        # 公司简介
        intro_text = f"""
{company_info.get('name', '该公司')}成立于{company_info.get('list_date', '上市')}，是一家专业从事
{company_info.get('industry', '相关')}行业的上市公司。公司于{company_info.get('list_date', '')}在
{company_info.get('market', '')}上市，股票代码为{self.project_info.get('stock_code', '')}。
"""
        elements.append({
            'type': 'paragraph',
            'content': intro_text.strip()
        })

        return elements

    def _generate_main_business_info(self) -> List[Dict[str, Any]]:
        """生成主营业务情况"""
        elements = []

        # 获取主营业务数据
        # TODO: 从Tushare获取主营业务构成数据
        elements.append({
            'type': 'paragraph',
            'content': '公司主营业务涵盖电力电缆、新材料等领域，主要产品包括电力电缆、光伏电缆、新材料产品等。'
        })

        return elements

    def _generate_dingzeng_info(self) -> List[Dict[str, Any]]:
        """生成定增项目概况"""
        elements = []

        # TODO: 从Tushare或配置文件获取定增相关信息
        dingzeng_info = [
            {'项目': '定增目的', '内容': '扩大产能、补充流动资金'},
            {'项目': '拟募集资金', '内容': '[待填入]'},
            {'项目': '发行价格', '内容': '[待填入]'},
            {'项目': '发行数量', '内容': '[待填入]'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['项目', '内容'],
            'data': dingzeng_info
        })

        return elements
