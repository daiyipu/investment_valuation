#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第四章：估值风险分析
包含DCF估值(三种方法)
"""

from typing import Dict, Any, List
import pandas as pd

from utils.dcf_calculator import DCFCalculator


class Chapter04Valuation:
    """估值风险分析章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any]):
        """初始化

        Args:
            config: 配置信息
            data: 报告数据
        """
        self.config = config
        self.data = data
        self.project_info = config.get('project', {})
        self.dcf_calculator = DCFCalculator()

    def generate(self) -> List[Dict[str, Any]]:
        """生成章节内容

        Returns:
            章节元素列表
        """
        elements = []

        # 章节标题
        elements.append({
            'type': 'heading',
            'content': '第四章 估值风险分析',
            'level': 1
        })

        # 4.1 当前估值水平
        elements.append({
            'type': 'heading',
            'content': '4.1 当前估值水平',
            'level': 2
        })

        elements.extend(self._generate_current_valuation())

        # 4.2 DCF估值分析
        elements.append({
            'type': 'heading',
            'content': '4.2 DCF估值分析',
            'level': 2
        })

        elements.extend(self._generate_dcf_valuation())

        # 4.3 历史估值对比
        elements.append({
            'type': 'heading',
            'content': '4.3 历史估值对比',
            'level': 2
        })

        elements.extend(self._generate_historical_valuation())

        # 4.4 估值风险评估
        elements.append({
            'type': 'heading',
            'content': '4.4 估值风险评估',
            'level': 2
        })

        elements.extend(self._generate_valuation_risk())

        return elements

    def _generate_current_valuation(self) -> List[Dict[str, Any]]:
        """生成当前估值水平"""
        elements = []

        valuation_data = self.data.get('valuation_data', {})

        # 估值指标表
        valuation_info = [
            {'指标': '最新收盘价', '数值': f"{valuation_data.get('current_price', 0):.2f}元" if valuation_data.get('current_price') else '[暂无数据]'},
            {'指标': '市盈率(PE)', '数值': f"{valuation_data.get('pe_ratio', 0):.2f}" if valuation_data.get('pe_ratio') else '[暂无数据]'},
            {'指标': '市净率(PB)', '数值': f"{valuation_data.get('pb_ratio', 0):.2f}" if valuation_data.get('pb_ratio') else '[暂无数据]'},
            {'指标': '总市值', '数值': f"{valuation_data.get('market_cap', 0)/10000:.2f}亿元" if valuation_data.get('market_cap') else '[暂无数据]'},
            {'指标': '流通市值', '数值': f"{valuation_data.get('float_market_cap', 0)/10000:.2f}亿元" if valuation_data.get('float_market_cap') else '[暂无数据]'},
        ]

        elements.append({
            'type': 'table',
            'headers': ['指标', '数值'],
            'data': valuation_info
        })

        return elements

    def _generate_dcf_valuation(self) -> List[Dict[str, Any]]:
        """生成DCF估值分析"""
        elements = []

        # DCF参数说明
        dcf_params_intro = """
DCF（折现现金流）估值采用以下三种方法进行计算：

1. FCFF法（企业价值）
   公式：EV = Σ FCFF/(1+WACC)^t + TV/(1+WACC)^n
   说明：使用加权平均资本成本(WACC)作为折现率，计算企业整体价值

2. FCFE法（股权价值）
   公式：股权价值 = Σ FCFE/(1+Ke)^t + TV/(1+Ke)^n
   说明：使用权益资本成本(Ke)作为折现率，直接计算股权价值

3. 间接法（股权价值）
   公式：股权价值 = EV - 净债务
   说明：先计算企业价值，再减去净债务得到股权价值

参数设置：
- 无风险利率：3%（参考10年期国债收益率）
- 市场风险溢价：6%
- Beta系数：1.0
- 永续增长率：3%
- 预测期：5年
"""
        elements.append({
            'type': 'paragraph',
            'content': dcf_params_intro.strip()
        })

        # 计算DCF估值
        financial_statements = self.data.get('financial_statements', {})
        financial_indicators = self.data.get('financial_indicators', pd.DataFrame())
        valuation_data = self.data.get('valuation_data', {})

        if not financial_statements.get('cashflow', pd.DataFrame()).empty:
            # 计算DCF估值
            dcf_result = self.dcf_calculator.calculate_dcf_valuation(
                financial_statements=financial_statements,
                financial_indicators=financial_indicators
            )

            # 获取总股本
            # 从配置或数据中获取
            shares_outstanding = self._get_shares_outstanding()
            current_price = valuation_data.get('current_price')

            # 生成估值汇总表
            valuation_summary = self.dcf_calculator.generate_valuation_summary(
                dcf_result, shares_outstanding, current_price
            )

            # DCF参数表
            dcf_params = [
                {'参数': 'WACC(加权平均资本成本)', '数值': f"{dcf_result.get('wacc', 0)*100:.2f}%"},
                {'参数': 'Ke(权益资本成本)', '数值': f"{dcf_result.get('ke', 0)*100:.2f}%"},
                {'参数': '无风险利率', '数值': f"{dcf_result.get('risk_free_rate', 0)*100:.2f}%"},
                {'参数': '市场风险溢价', '数值': f"{dcf_result.get('market_risk_premium', 0)*100:.2f}%"},
                {'参数': 'Beta系数', '数值': f"{dcf_result.get('beta', 0):.2f}"},
                {'参数': '永续增长率', '数值': f"{dcf_result.get('terminal_growth_rate', 0)*100:.2f}%"},
            ]

            elements.append({
                'type': 'heading',
                'content': '4.2.1 DCF参数设置',
                'level': 3
            })

            elements.append({
                'type': 'table',
                'headers': ['参数', '数值'],
                'data': dcf_params
            })

            # 估值结果表
            elements.append({
                'type': 'heading',
                'content': '4.2.2 DCF估值结果',
                'level': 3
            })

            # 转换DataFrame为列表格式
            valuation_rows = []
            for _, row in valuation_summary.iterrows():
                row_dict = {
                    '估值方法': row['估值方法'],
                    '估值结果(万元)': f"{row['估值结果(万元)']:.2f}",
                }
                if '每股价值(元)' in row and pd.notna(row.get('每股价值(元)')):
                    row_dict['每股价值(元)'] = f"{row['每股价值(元)']:.2f}"
                else:
                    row_dict['每股价值(元)'] = '[待计算]'
                if '估值空间' in row and pd.notna(row.get('估值空间')):
                    row_dict['估值空间'] = row['估值空间']
                else:
                    row_dict['估值空间'] = '-'
                row_dict['说明'] = row.get('说明', '')
                valuation_rows.append(row_dict)

            elements.append({
                'type': 'table',
                'headers': ['估值方法', '估值结果(万元)', '每股价值(元)', '估值空间', '说明'],
                'data': valuation_rows
            })

            # 三种方法对比分析
            elements.append({
                'type': 'heading',
                'content': '4.2.3 三种方法对比分析',
                'level': 3
            })

            method1_ev = dcf_result.get('method1_ev', 0)
            method2_equity = dcf_result.get('method2_equity_value', 0)
            method3_equity = dcf_result.get('method3_equity_value', 0)

            # 计算差异
            if method2_equity > 0 and method3_equity > 0:
                diff_2_3 = abs(method2_equity - method3_equity) / method2_equity * 100
            else:
                diff_2_3 = 0

            comparison_text = f"""
三种DCF估值方法对比：

方法一(FCFF法)计算结果：企业价值 {method1_ev/10000:.2f} 亿元
方法二(FCFE法)计算结果：股权价值 {method2_equity/10000:.2f} 亿元
方法三(间接法)计算结果：股权价值 {method3_equity/10000:.2f} 亿元

三种方法的理论一致性：FCFE = EV - 净债务
方法二与方法三的差异：{diff_2_3:.2f}%

差异原因分析：
- FCFF法和FCFE法使用不同的折现率(WACC vs Ke)
- 间接法通过企业价值减去净债务得到股权价值
- 三种方法相互验证，可综合判断估值合理性
"""
            elements.append({
                'type': 'paragraph',
                'content': comparison_text.strip()
            })

        else:
            elements.append({
                'type': 'paragraph',
                'content': '[现金流数据不足，暂无法进行DCF估值分析]'
            })

        return elements

    def _get_shares_outstanding(self) -> float:
        """获取总股本

        Returns:
            总股本(万股)
        """
        # 从财务数据中尝试获取
        balance_sheet = self.data.get('financial_statements', {}).get('balance_sheet', pd.DataFrame())

        if not balance_sheet.empty:
            # 尝试查找股本相关字段
            for col in balance_sheet.columns:
                if '股本' in str(col):
                    value = balance_sheet[col].iloc[-1] if len(balance_sheet) > 0 else None
                    if value and pd.notna(value):
                        return float(value)

        # 默认值
        return 100000  # 10亿股 = 100000万股

    def _generate_historical_valuation(self) -> List[Dict[str, Any]]:
        """生成历史估值对比"""
        elements = []

        # TODO: 从Tushare获取历史PE/PB数据
        elements.append({
            'type': 'paragraph',
            'content': '[历史估值数据待从Tushare获取分析]'
        })

        return elements

    def _generate_valuation_risk(self) -> List[Dict[str, Any]]:
        """生成估值风险评估"""
        elements = []

        valuation_data = self.data.get('valuation_data', {})
        pe_ratio = valuation_data.get('pe_ratio', 0)

        # DCF估值结果
        financial_statements = self.data.get('financial_statements', {})
        dcf_result = None
        if not financial_statements.get('cashflow', pd.DataFrame()).empty:
            dcf_result = self.dcf_calculator.calculate_dcf_valuation(
                financial_statements=financial_statements,
                financial_indicators=self.data.get('financial_indicators', pd.DataFrame())
            )

        # 估值风险评估
        if pe_ratio:
            if pe_ratio < 20:
                risk_level = '低'
                risk_desc = '估值处于较低水平，具有一定投资价值'
            elif pe_ratio < 40:
                risk_level = '中'
                risk_desc = '估值处于合理区间'
            else:
                risk_level = '较高'
                risk_desc = '估值偏高，需关注'
        else:
            risk_level = '待评估'
            risk_desc = '暂无估值数据'

        risk_evaluation = [
            {'评估维度': '当前PE估值', '评估结果': f"{pe_ratio:.2f}" if pe_ratio else '[暂无数据]', '风险等级': risk_level},
            {'评估维度': 'DCF估值', '评估结果': f"{dcf_result.get('method2_equity_value', 0)/10000:.2f}亿元" if dcf_result else '[暂无数据]', '风险等级': '参考'},
        ]

        # 如果有当前市值和DCF估值，计算估值偏离度
        if dcf_result and valuation_data.get('market_cap'):
            market_cap = valuation_data.get('market_cap', 0) / 10000  # 转换为亿元
            dcf_value = dcf_result.get('method3_equity_value', 0) / 10000  # 转换为亿元

            if dcf_value > 0:
                deviation = (market_cap - dcf_value) / dcf_value * 100
                deviation_level = '低估' if deviation < -10 else ('高估' if deviation > 10 else '合理')
                risk_evaluation.append({
                    '评估维度': '市值/DCF偏离度',
                    '评估结果': f"当前市值{market_cap:.2f}亿 vs DCF {dcf_value:.2f}亿 ({deviation:+.2f}%)",
                    '风险等级': deviation_level
                })

        elements.append({
            'type': 'table',
            'headers': ['评估维度', '评估结果', '风险等级'],
            'data': risk_evaluation
        })

        return elements
