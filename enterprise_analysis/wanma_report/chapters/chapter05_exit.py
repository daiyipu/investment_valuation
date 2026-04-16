#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第五章：退出风险分析
"""

from typing import Dict, Any, List


class Chapter05Exit:
    """退出风险分析章节生成器"""

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
            'content': '第五章 退出风险分析',
            'level': 1
        })

        # 5.1 退出方式分析
        elements.append({
            'type': 'heading',
            'content': '5.1 退出方式分析',
            'level': 2
        })

        elements.extend(self._generate_exit_methods())

        # 5.2 退出渠道分析
        elements.append({
            'type': 'heading',
            'content': '5.2 退出渠道分析',
            'level': 2
        })

        elements.extend(self._generate_exit_channels())

        # 5.3 退出风险评估
        elements.append({
            'type': 'heading',
            'content': '5.3 退出风险评估',
            'level': 2
        })

        elements.extend(self._generate_exit_risk())

        return elements

    def _generate_exit_methods(self) -> List[Dict[str, Any]]:
        """生成退出方式分析"""
        elements = []

        # 退出方式对比表
        exit_methods = [
            {'退出方式': '二级市场减持', '优点': '灵活、变现快', '缺点': '受减持规则限制', '适用性': '主要退出方式'},
            {'退出方式': '大宗交易', '优点': '交易量大、价格稳定', '缺点': '受接盘方限制', '适用性': '适合大额减持'},
            {'退出方式': '协议转让', '优点': '可一次性转让', '缺点': '需找到合适受让方', '适用性': '适合战略转让'},
            {'退出方式': '定增锁定期满', '优点': '可批量减持', '缺点': '锁定期较长', '适用性': '定增项目特有'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['退出方式', '优点', '缺点', '适用性'],
            'data': exit_methods
        })

        return elements

    def _generate_exit_channels(self) -> List[Dict[str, Any]]:
        """生成退出渠道分析"""
        elements = []

        company_info = self.data.get('company_info', {})
        market = company_info.get('market', '')

        # 退出渠道分析
        channels = [
            {'渠道类型': 'A股主板', '适用性': '适用' if '主板' in str(market) or market in ['SH', 'SZ'] else '不适用', '说明': '可通过二级市场减持'},
            {'渠道类型': '科创板/创业板', '适用性': '适用' if '科创' in str(market) or '创业' in str(market) else '不适用', '说明': '减持规则有所差异'},
            {'渠道类型': '大宗交易平台', '适用性': '适用', '说明': '适合大额减持'},
            {'渠道类型': '协议转让', '适用性': '适用', '说明': '需符合相关条件'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['渠道类型', '适用性', '说明'],
            'data': channels
        })

        return elements

    def _generate_exit_risk(self) -> List[Dict[str, Any]]:
        """生成退出风险评估"""
        elements = []

        valuation_data = self.data.get('valuation_data', {})
        scoring_result = self.data.get('scoring_result', {})

        # 退出风险评估
        pe_ratio = valuation_data.get('pe_ratio')
        exit_risks = [
            {'风险类型': '锁定期风险', '评估': '定增项目存在锁定期，需等待锁定期满', '等级': '中'},
            {'风险类型': '市场风险', '评估': '减持时点市场环境存在不确定性', '等级': '中'},
            {'风险类型': '流动性风险', '评估': '大额减持可能影响股价', '等级': '中'},
            {'风险类型': '价格风险', '评估': f"当前PE: {pe_ratio:.2f}" if pe_ratio else '暂无PE数据', '等级': '待评估'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['风险类型', '评估', '等级'],
            'data': exit_risks
        })

        return elements
