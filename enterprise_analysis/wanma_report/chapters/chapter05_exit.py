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

        valuation_data = self.data.get('valuation_data', {})
        market_cap = valuation_data.get('market_cap')
        float_market_cap = valuation_data.get('float_market_cap')

        # 计算减持影响
        liquidity_note = ''
        if market_cap and float_market_cap and float_market_cap > 0:
            float_ratio = float_market_cap / market_cap * 100
            liquidity_note = f'当前流通市值{float_market_cap / 10000:.2f}亿元，占总市值{float_ratio:.1f}%'

        exit_methods = [
            {'退出方式': '二级市场减持', '优点': '灵活、变现快', '缺点': '受减持规则限制(每3个月不超过1%)', '适用性': '主要退出方式'},
            {'退出方式': '大宗交易', '优点': '交易量大、对市场冲击小', '缺点': '需寻找接盘方，可能有折价', '适用性': '适合大额减持'},
            {'退出方式': '协议转让', '优点': '可一次性转让较大比例', '缺点': '需找到合适受让方，审批流程较长', '适用性': '适合战略转让'},
            {'退出方式': '定增锁定期满后减持', '优点': '定增股份锁定期结束后可灵活安排', '缺点': '锁定期较长(通常6-18个月)', '适用性': '定增项目特有'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['退出方式', '优点', '缺点', '适用性'],
            'data': exit_methods
        })

        if liquidity_note:
            elements.append({
                'type': 'paragraph',
                'content': liquidity_note
            })

        return elements

    def _generate_exit_channels(self) -> List[Dict[str, Any]]:
        """生成退出渠道分析"""
        elements = []

        company_info = self.data.get('company_info', {})
        market = company_info.get('market', '')
        valuation_data = self.data.get('valuation_data', {})
        turnover_rate = valuation_data.get('turnover_rate')

        # 判断上市板块
        board = '中小板'
        if '主板' in str(market):
            board = '主板'
        elif '科创' in str(market):
            board = '科创板'
        elif '创业' in str(market):
            board = '创业板'

        # 流动性评估
        liquidity = '一般'
        if turnover_rate:
            if turnover_rate > 3:
                liquidity = '较好'
            elif turnover_rate > 1:
                liquidity = '一般'
            else:
                liquidity = '较差'

        channels = [
            {'渠道类型': f'A股{board}', '适用性': '适用', '说明': f'公司上市于{board}，可通过二级市场正常减持'},
            {'渠道类型': '大宗交易平台', '适用性': '适用', '说明': '适合单笔大额减持，降低对二级市场冲击'},
            {'渠道类型': '协议转让', '适用性': '适用', '说明': '需符合5%以上股东减持相关规定'},
            {'渠道类型': '非交易过户', '适用性': '适用', '说明': '适用于特定情形下的股份转移'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['渠道类型', '适用性', '说明'],
            'data': channels
        })

        # 流动性分析
        liquidity_text = f'当前日均换手率{turnover_rate:.2f}%' if turnover_rate else '换手率数据暂缺'
        liquidity_text += f'，市场流动性{liquidity}。'
        if liquidity == '较差':
            liquidity_text += '需注意大额减持可能面临流动性不足风险，建议分批减持。'

        elements.append({
            'type': 'paragraph',
            'content': liquidity_text
        })

        return elements

    def _generate_exit_risk(self) -> List[Dict[str, Any]]:
        """生成退出风险评估"""
        elements = []

        valuation_data = self.data.get('valuation_data', {})
        scoring_result = self.data.get('scoring_result', {})
        company_info = self.data.get('company_info', {})

        pe_ratio = valuation_data.get('pe_ratio')
        pb_ratio = valuation_data.get('pb_ratio')
        turnover_rate = valuation_data.get('turnover_rate')
        total_score = scoring_result.get('total_score', 0)

        # 锁定期风险评估
        lockup_risk = '中'
        lockup_desc = '定增项目存在法定锁定期(通常6个月)，锁定期内无法减持'

        # 市场风险评估
        market_risk = '中'
        if pe_ratio:
            if pe_ratio > 40:
                market_risk = '较高'
                market_desc = f'当前PE({pe_ratio:.1f})处于较高水平，未来市场估值回调风险较大'
            elif pe_ratio < 15:
                market_risk = '低'
                market_desc = f'当前PE({pe_ratio:.1f})处于较低水平，市场估值风险较小'
            else:
                market_desc = f'当前PE({pe_ratio:.1f})处于合理区间，市场风险适中'
        else:
            market_desc = '暂无PE数据，市场风险难以评估'

        # 流动性风险评估
        if turnover_rate:
            if turnover_rate < 1:
                liquidity_risk = '较高'
                liquidity_desc = f'日均换手率仅{turnover_rate:.2f}%，流动性不足，大额减持可能面临困难'
            elif turnover_rate < 3:
                liquidity_risk = '中'
                liquidity_desc = f'日均换手率{turnover_rate:.2f}%，流动性一般'
            else:
                liquidity_risk = '低'
                liquidity_desc = f'日均换手率{turnover_rate:.2f}%，流动性良好'
        else:
            liquidity_risk = '待评估'
            liquidity_desc = '暂无换手率数据'

        # 价格风险评估
        if pb_ratio and pb_ratio < 1:
            price_risk = '较高'
            price_desc = f'当前PB({pb_ratio:.2f})低于1，股价已破净，需关注基本面恶化风险'
        elif pe_ratio:
            price_risk = '中'
            price_desc = f'当前PE: {pe_ratio:.2f}' + (f'，PB: {pb_ratio:.2f}' if pb_ratio else '')
        else:
            price_risk = '待评估'
            price_desc = '暂无估值数据'

        exit_risks = [
            {'风险类型': '锁定期风险', '评估': lockup_desc, '等级': lockup_risk},
            {'风险类型': '市场风险', '评估': market_desc, '等级': market_risk},
            {'风险类型': '流动性风险', '评估': liquidity_desc, '等级': liquidity_risk},
            {'风险类型': '价格风险', '评估': price_desc, '等级': price_risk},
        ]

        elements.append({
            'type': 'table',
            'headers': ['风险类型', '评估', '等级'],
            'data': exit_risks
        })

        # 退出建议
        suggestions = []
        if total_score >= 70:
            suggestions.append('公司基本面较好，可在锁定期结束后择机减持')
        else:
            suggestions.append('公司基本面存在风险，建议制定分批退出计划')

        if turnover_rate and turnover_rate < 1:
            suggestions.append('流动性较差，建议以大宗交易为主，减少对二级市场冲击')

        if pe_ratio and pe_ratio > 40:
            suggestions.append('估值偏高，建议不要集中减持，分散在多个时间窗口')

        if suggestions:
            elements.append({
                'type': 'paragraph',
                'content': '退出建议：' + '；'.join(suggestions) + '。'
            })

        return elements
