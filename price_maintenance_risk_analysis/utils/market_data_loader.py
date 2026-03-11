# -*- coding: utf-8 -*-
"""
市场数据加载工具

功能：
- 从JSON文件加载市场数据（由update_market_data.py生成）
- 提供风险参数计算和数据摘要输出
"""

import json
import os
from typing import Dict, Optional


def load_market_data(stock_code: str, data_dir: str = 'data') -> Optional[Dict]:
    """
    加载股票的市场数据（波动率、收益率等）

    参数:
        stock_code: 股票代码，如 '300735.SZ' 或 '300735_SZ'
        data_dir: 数据文件所在目录，默认为'data'目录

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


def load_market_indices_data(data_dir: str = '..') -> Optional[Dict]:
    """
    加载市场指数数据（波动率、收益率等）

    参数:
        data_dir: 数据文件所在目录，默认为上级目录

    返回:
        包含指数数据的字典，如果文件不存在则返回None

    使用示例:
        >>> indices_data = load_market_indices_data()
        >>> if indices_data:
        ...     hs300_data = indices_data['沪深300']
        ...     print(f"沪深300波动率: {hs300_data['volatility']*100:.2f}%")
        ...     print(f"沪深300收益率: {hs300_data['drift']*100:.2f}%")
    """
    filepath = os.path.join(data_dir, 'market_indices_scenario_data.json')

    if not os.path.exists(filepath):
        print(f"⚠️ 指数数据文件不存在: {filepath}")
        print(f"   提示: 请先运行 07_market_data_analysis.ipynb 第二部分生成数据")
        return None

    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            indices_data = json.load(f)

        print(f"✅ 已加载指数数据: {filepath}")
        print(f"   指数数量: {len(indices_data)}")

        return indices_data

    except Exception as e:
        print(f"❌ 加载指数数据失败: {e}")
        return None


def print_indices_data_summary(indices_data: Dict):
    """
    打印指数数据摘要
    """
    if not indices_data:
        print("⚠️ 无指数数据")
        return

    print("\n" + "="*70)
    print(f"📊 市场指数数据摘要")
    print("="*70)

    print(f"\n{'指数':<12} {'当前点位':<15} {'波动率(60日)':<15} {'收益率(60日)':<15}")
    print('-'*60)

    for name, data in indices_data.items():
        print(f"{name:<12} {data.get('current_level', 0):>10.2f}       "
              f"{data.get('volatility', 0)*100:>10.2f}%       "
              f"{data.get('drift', 0)*100:>10.2f}%")

    print("="*70)
    print(f"\n📌 可用指数:")
    for name in indices_data.keys():
        print(f"   - {name}")
    print(f"\n💡 使用示例:")
    print(f"   indices_data = load_market_indices_data()")
    print(f"   volatility = indices_data['沪深300']['volatility']")
    print(f"   drift = indices_data['沪深300']['drift']")


def get_index_risk_params(indices_data: Dict, index_name: str = '沪深300',
                          volatility_window: str = '60d') -> Dict:
    """
    从指数数据中提取指定指数的风险参数

    参数:
        indices_data: load_market_indices_data()返回的数据
        index_name: 指数名称 ('沪深300', '上证指数', 等)
        volatility_window: 波动率窗口 ('30d', '60d', '120d', '180d')

    返回:
        风险参数字典
    """
    if not indices_data or index_name not in indices_data:
        print(f"⚠️ 指数 {index_name} 数据不存在")
        return {
            'volatility': 0.20,  # 指数默认波动率较低
            'drift': 0.05,
        }

    index_data = indices_data[index_name]

    return {
        'volatility': index_data.get(f'volatility_{volatility_window}',
                                     index_data.get('volatility', 0.20)),
        'drift': index_data.get(f'return_{volatility_window}',
                                index_data.get('drift', 0.05)),
        'current_level': index_data.get('current_level', 0),
    }


def save_market_indices_data(indices_results: Dict, data_dir: str = '..'):
    """
    保存市场指数分析结果到JSON文件（统一的数据导出接口）

    参数:
        indices_results: analyze_all_indices()或批量分析返回的结果字典
        data_dir: 数据保存目录

    使用示例:
        >>> from utils.market_data_loader import save_market_indices_data
        >>> # 在notebook中分析完指数后
        >>> save_market_indices_data(all_indices_results)
    """
    # 创建简化版数据结构，供情景分析使用
    scenario_indices_data = {}

    for index_name, result in indices_results.items():
        scenario_indices_data[index_name] = {
            'code': result['index_code'],
            'name': index_name,
            'current_level': result['current_close'],

            # 波动率（多个窗口）
            'volatility_30d': result['volatility']['30日']['latest'],
            'volatility_60d': result['volatility']['60日']['latest'],
            'volatility_120d': result['volatility']['120日']['latest'],
            'volatility_180d': result['volatility']['180日']['latest'],
            'volatility_250d': result['volatility'].get('250日', {}).get('latest', result['volatility']['180日']['latest']),
            'volatility': result['volatility']['60日']['latest'],  # 默认60日

            # 收益率（多个窗口）
            'return_30d': result['returns']['30日']['annualized_return'],
            'return_60d': result['returns']['60日']['annualized_return'],
            'return_120d': result['returns']['120日']['annualized_return'],
            'return_180d': result['returns']['180日']['annualized_return'],
            'return_250d': result['returns'].get('250日', {}).get('annualized_return', result['returns']['180日']['annualized_return']),
            'drift': result['returns']['60日']['annualized_return'],  # 默认60日

            # 技术指标
            'ma_30': result['price']['30日']['latest_ma'],
            'ma_60': result['price']['60日']['latest_ma'],
            'ma_120': result['price']['120日']['latest_ma'],
            'ma_180': result['price']['180日']['latest_ma'],
            'ma_250': result['price'].get('250日', {}).get('latest_ma', result['price']['180日']['latest_ma']),

            # 胜率（多个窗口）
            'win_rate_60d': result['returns']['60日']['win_rate'],
            'win_rate_120d': result['returns']['120日']['win_rate'],
            'win_rate_250d': result['returns'].get('250日', {}).get('win_rate', result['returns']['180日']['win_rate']),
        }

    # 保存为JSON文件
    output_file = os.path.join(data_dir, 'market_indices_scenario_data.json')
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(scenario_indices_data, f, indent=2, ensure_ascii=False)

    print(f"💾 指数数据已保存: {output_file}")
    print(f"   指数数量: {len(scenario_indices_data)}")


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
