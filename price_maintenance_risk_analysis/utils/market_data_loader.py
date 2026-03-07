# -*- coding: utf-8 -*-
"""
市场数据加载工具
从07_market_data_analysis生成的JSON文件加载真实市场数据
"""

import json
import os
from typing import Dict, Optional


def load_market_data(stock_code: str, data_dir: str = '..') -> Optional[Dict]:
    """
    加载股票的市场数据（波动率、收益率等）

    参数:
        stock_code: 股票代码，如 '300735.SZ' 或 '300735_SZ'
        data_dir: 数据文件所在目录，默认为上级目录

    返回:
        包含市场数据的字典，如果文件不存在则返回None

    使用示例:
        >>> market_data = load_market_data('300735.SZ')
        >>> if market_data:
        ...     volatility = market_data['volatility']
        ...     drift = market_data['drift']
        ...     print(f"波动率: {volatility*100:.2f}%")
        ...     print(f"收益率: {drift*100:.2f}%")
    """
    # 统一文件名格式（将点替换为下划线）
    filename = stock_code.replace('.', '_') + '_market_data.json'
    filepath = os.path.join(data_dir, filename)

    if not os.path.exists(filepath):
        print(f"⚠️ 市场数据文件不存在: {filepath}")
        print(f"   提示: 请先运行 07_market_data_analysis.ipynb 生成数据")
        return None

    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            market_data = json.load(f)

        print(f"✅ 已加载市场数据: {filepath}")
        print(f"   股票: {market_data.get('stock_name', 'N/A')} ({market_data.get('stock_code', 'N/A')})")
        print(f"   分析日期: {market_data.get('analysis_date', 'N/A')}")
        print(f"   当前价格: {market_data.get('current_price', 0):.2f} 元")

        return market_data

    except Exception as e:
        print(f"❌ 加载市场数据失败: {e}")
        return None


def get_risk_params(market_data: Dict, volatility_window: str = '60d') -> Dict:
    """
    从市场数据中提取风险参数

    参数:
        market_data: load_market_data()返回的数据
        volatility_window: 波动率窗口 ('30d', '60d', '120d', '180d')

    返回:
        风险参数字典
    """
    if not market_data:
        return {
            'volatility': 0.35,  # 默认值
            'drift': 0.08,       # 默认值
        }

    return {
        'volatility': market_data.get(f'volatility_{volatility_window}',
                                      market_data.get('volatility', 0.35)),
        'drift': market_data.get(f'annual_return_{volatility_window}',
                                 market_data.get('drift', 0.08)),
    }


def print_market_data_summary(market_data: Dict):
    """
    打印市场数据摘要
    """
    if not market_data:
        print("⚠️ 无市场数据")
        return

    print("\n" + "="*70)
    print(f"📊 {market_data.get('stock_name', 'N/A')} 市场数据摘要")
    print("="*70)

    print(f"\n📈 价格数据:")
    print(f"   当前价格: {market_data.get('current_price', 0):.2f} 元")
    print(f"   平均价格: {market_data.get('avg_price_all', 0):.2f} 元")
    print(f"   价格中位数: {market_data.get('median_price', 0):.2f} 元")

    print(f"\n⚠️ 波动率:")
    print(f"   30日波动率: {market_data.get('volatility_30d', 0)*100:.2f}%")
    print(f"   60日波动率: {market_data.get('volatility_60d', 0)*100:.2f}%")
    print(f"   120日波动率: {market_data.get('volatility_120d', 0)*100:.2f}%")
    print(f"   180日波动率: {market_data.get('volatility_180d', 0)*100:.2f}%")

    print(f"\n📊 收益率:")
    print(f"   30日年化收益: {market_data.get('annual_return_30d', 0)*100:.2f}%")
    print(f"   60日年化收益: {market_data.get('annual_return_60d', 0)*100:.2f}%")
    print(f"   120日年化收益: {market_data.get('annual_return_120d', 0)*100:.2f}%")
    print(f"   180日年化收益: {market_data.get('annual_return_180d', 0)*100:.2f}%")

    print(f"\n📈 移动平均线:")
    print(f"   MA30: {market_data.get('ma_30', 0):.2f} 元")
    print(f"   MA60: {market_data.get('ma_60', 0):.2f} 元")
    print(f"   MA120: {market_data.get('ma_120', 0):.2f} 元")
    print(f"   MA180: {market_data.get('ma_180', 0):.2f} 元")

    print(f"\n📊 其他统计:")
    print(f"   60日胜率: {market_data.get('win_rate_60d', 0)*100:.1f}%")
    print(f"   数据天数: {market_data.get('total_days', 0)} 天")

    print("="*70)


# 使用示例
if __name__ == '__main__':
    # 测试加载光弘科技的市场数据
    data = load_market_data('300735.SZ')

    if data:
        # 打印摘要
        print_market_data_summary(data)

        # 获取风险参数
        risk_params = get_risk_params(data, volatility_window='60d')
        print(f"\n推荐风险参数:")
        print(f"  波动率: {risk_params['volatility']*100:.2f}%")
        print(f"  漂移率: {risk_params['drift']*100:.2f}%")
