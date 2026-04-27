#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
市场风险溢价计算模块
根据历史数据计算中国A股市场的风险溢价（Rm - Rf）

说明：
1. 市场风险溢价是宏观经济参数，不应该因行业而异
2. 行业风险差异通过Beta系数体现，而非市场风险溢价
3. 推荐使用投行标准值（6-8%）作为保守估计
"""

import numpy as np
import pandas as pd
from typing import Dict, Optional
from datetime import datetime, timedelta


class MarketRiskPremiumCalculator:
    """市场风险溢价计算器"""

    # 中国A股市场参考值
    CHINA_MARKET_REFERENCES = {
        'conservative': 0.08,      # 保守估计（投行上限）
        'standard': 0.07,          # 标准值（投行推荐）
        'moderate': 0.06,          # 中等值（投行下限）
        'historical_10y': 0.0125,  # 历史10年平均（较低）
    }

    # 全球市场参考值（Damodaran等权威研究）
    GLOBAL_REFERENCES = {
        'US': 0.045,               # 美国（成熟市场）
        'China': 0.07,             # 中国（新兴市场）
        'Emerging_Markets': 0.075, # 新兴市场平均
        'Global': 0.055,           # 全球平均
    }

    def __init__(self, pro_api=None):
        """
        初始化计算器

        参数:
            pro_api: tushare pro_api对象
        """
        self.pro = pro_api

    def calculate_from_history(
        self,
        index_code: str = '000300.SH',
        years: int = 10,
        method: str = 'geometric'
    ) -> Dict:
        """
        从历史数据计算市场风险溢价

        参数:
            index_code: 指数代码（默认沪深300）
            years: 历史年数
            method: 计算方法 ('geometric'几何平均 或 'arithmetic'算术平均)

        返回:
            {
                'market_risk_premium': float,  # 市场风险溢价
                'market_return': float,        # 市场年化收益率
                'risk_free_rate': float,       # 无风险利率
                'data_points': int,
                'calculation_method': str
            }
        """
        if self.pro is None:
            return {
                'market_risk_premium': 0.07,
                'error': 'API未初始化，使用标准值7%'
            }

        try:
            # 获取历史数据
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=years*365)).strftime('%Y%m%d')

            df_index = self.pro.index_daily(
                ts_code=index_code,
                start_date=start_date,
                end_date=end_date,
                fields='trade_date,close,pct_chg'
            )

            if df_index.empty:
                return {'market_risk_premium': 0.07, 'error': '未获取到数据'}

            df_index = df_index.sort_values('trade_date').reset_index(drop=True)

            # 计算年化收益率
            first_price = df_index['close'].iloc[0]
            last_price = df_index['close'].iloc[-1]
            total_return = (last_price / first_price) - 1
            n_days = len(df_index)

            if method == 'geometric':
                # 几何平均（更准确）
                market_return = (1 + total_return) ** (252 / n_days) - 1
            else:
                # 算术平均
                daily_return_mean = df_index['pct_chg'].mean() / 100
                market_return = daily_return_mean * 252

            # 无风险利率（10年期国债）
            # 实际应用中应从国债数据获取，这里使用固定值
            risk_free_rate = 0.0215  # 约2.15%

            # 市场风险溢价
            mrp = market_return - risk_free_rate

            return {
                'market_risk_premium': round(mrp, 4),
                'market_return': round(market_return, 4),
                'risk_free_rate': risk_free_rate,
                'data_points': n_days,
                'calculation_method': f'{method}_average_{years}y',
                'index_code': index_code
            }

        except Exception as e:
            print(f"⚠️ 计算失败: {e}")
            return {'market_risk_premium': 0.07, 'error': str(e)}

    def get_recommended_premium(
        self,
        approach: str = 'standard'
    ) -> Dict:
        """
        获取推荐的市场风险溢价

        参数:
            approach: 选择方法
                - 'conservative': 保守值（8%）
                - 'standard': 标准值（7%）【推荐】
                - 'moderate': 中等值（6%）
                - 'historical': 历史计算值（需要API）

        返回:
            {
                'market_risk_premium': float,
                'approach': str,
                'rationale': str
            }
        """
        if approach == 'conservative':
            return {
                'market_risk_premium': self.CHINA_MARKET_REFERENCES['conservative'],
                'approach': 'conservative',
                'rationale': '保守估计（8%），适用于高不确定性环境，提供更高安全边际'
            }
        elif approach == 'moderate':
            return {
                'market_risk_premium': self.CHINA_MARKET_REFERENCES['moderate'],
                'approach': 'moderate',
                'rationale': '中等估计（6%），投行下限，适用于市场乐观预期'
            }
        elif approach == 'historical':
            if self.pro:
                result = self.calculate_from_history(years=10)
                mrp = result.get('market_risk_premium', 0.07)
                return {
                    'market_risk_premium': mrp,
                    'approach': 'historical_10y',
                    'rationale': f'基于沪深300过去10年历史数据计算（{mrp*100:.2f}%）'
                }
            else:
                return {
                    'market_risk_premium': 0.07,
                    'approach': 'standard',
                    'rationale': 'API未初始化，使用标准值7%'
                }
        else:  # 'standard' (default)
            return {
                'market_risk_premium': self.CHINA_MARKET_REFERENCES['standard'],
                'approach': 'standard',
                'rationale': '标准值（7%），投行推荐值，适用于大多数DCF估值场景'
            }


def get_market_risk_premium(
    pro_api=None,
    approach: str = 'standard'
) -> float:
    """
    快速获取市场风险溢价

    参数:
        pro_api: tushare API对象
        approach: 计算方法（'conservative', 'standard', 'moderate', 'historical'）

    返回:
        市场风险溢价（小数形式）

    使用示例:
        >>> mrp = get_market_risk_premium(approach='standard')
        >>> print(f"市场风险溢价: {mrp*100:.1f}%")
    """
    calculator = MarketRiskPremiumCalculator(pro_api)
    result = calculator.get_recommended_premium(approach)
    return result['market_risk_premium']


if __name__ == '__main__':
    print("📊 市场风险溢价计算器\n")

    # 测试不同方法
    approaches = ['conservative', 'standard', 'moderate']
    calculator = MarketRiskPremiumCalculator()

    print("=" * 70)
    print("中国A股市场风险溢价参考值")
    print("=" * 70)

    for approach in approaches:
        result = calculator.get_recommended_premium(approach)
        mrp = result['market_risk_premium']
        rationale = result['rationale']
        print(f"\n{approach.upper()}:")
        print(f"   数值: {mrp*100:.1f}%")
        print(f"   说明: {rationale}")

    # 如果有API，计算历史值
    try:
        import tushare as ts
        import os
        token = os.environ.get('TUSHARE_TOKEN', '')
        if token:
            pro = ts.pro_api(token)
            calculator_with_api = MarketRiskPremiumCalculator(pro_api=pro)

            print(f"\n{'=' * 70}")
            print("基于历史数据计算")
            print("=" * 70)

            for years in [3, 5, 10]:
                result = calculator_with_api.calculate_from_history(years=years)
                mrp = result['market_risk_premium']
                market_return = result['market_return']
                print(f"\n{years}年历史:")
                print(f"   市场年化收益: {market_return*100:.2f}%")
                print(f"   市场风险溢价: {mrp*100:.2f}%")
    except ImportError:
        print("\n⚠️ 未安装tushare，跳过历史数据计算")

    print(f"\n{'=' * 70}")
    print("✅ 推荐：使用标准值 7.0%")
    print("=" * 70)
    print("\n说明：")
    print("1. 市场风险溢价是宏观经济参数，不应因行业而异")
    print("2. 行业风险差异通过Beta系数体现（β > 1 为高Beta行业）")
    print("3. CAPM公式：Re = Rf + β × (Rm - Rf)")
    print("4. 例如：电子制造行业β≈1.7，其股权成本 = 3% + 1.7 × 7% = 14.9%")
    print("5. 如果使用β=1.0的行业标准，则股权成本 = 3% + 1.0 × 7% = 10%")
    print("   → 行业差异已通过β体现，无需调整市场风险溢价")
