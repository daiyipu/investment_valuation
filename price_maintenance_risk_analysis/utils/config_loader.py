# -*- coding: utf-8 -*-
"""
定增分析配置加载工具
统一加载定增参数和市场数据，供所有脚本使用
"""

import json
import os
from typing import Dict, Optional, Tuple


def load_placement_config(
    stock_code: str,
    data_dir: str = None,
    prefer_market_data: bool = True
) -> Tuple[Dict, Dict, Dict]:
    """
    加载定增分析的完整配置（定增参数 + 市场数据）

    参数:
        stock_code: 股票代码，如 '300735.SZ'
        data_dir: 数据文件所在目录（None时自动检测）
        prefer_market_data: 是否优先使用市场数据（如果存在）

    返回:
        (project_params, risk_params, market_data) 元组
        - project_params: 定增项目参数
        - risk_params: 风险参数（波动率、收益率等）
        - market_data: 市场数据（如果有）

    使用示例:
        >>> project_params, risk_params, market_data = load_placement_config('300735.SZ')
        >>> print(f"波动率: {risk_params['volatility']*100:.2f}%")
        >>> print(f"收益率: {risk_params['drift']*100:.2f}%")
    """
    # 自动检测数据目录
    if data_dir is None:
        # 尝试多个可能的路径
        possible_dirs = ['.', '..', 'data', '../data']
        for test_dir in possible_dirs:
            test_file = os.path.join(test_dir, f"{stock_code.replace('.', '_')}_placement_params.json")
            if os.path.exists(test_file):
                data_dir = test_dir
                break

        if data_dir is None:
            data_dir = '.'  # 默认当前目录

    # 统一文件名格式
    placement_file = os.path.join(data_dir, f"{stock_code.replace('.', '_')}_placement_params.json")

    # 1. 加载定增参数
    if not os.path.exists(placement_file):
        raise FileNotFoundError(f"定增参数文件不存在: {placement_file}")

    with open(placement_file, 'r', encoding='utf-8') as f:
        placement_params = json.load(f)

    print(f"✅ 已加载定增参数: {placement_file}")

    # 2. 尝试加载市场数据（如果启用）
    market_data = None
    if prefer_market_data:
        try:
            from utils.market_data_loader import load_market_data
            market_data = load_market_data(stock_code, data_dir=data_dir)
        except Exception as e:
            print(f"⚠️ 市场数据加载失败: {e}")
            print(f"   将使用定增参数文件中的数据")

    # 3. 构建项目参数
    project_params = {
        'issue_price': placement_params['issue_price'],
        'issue_shares': placement_params['issue_shares'],
        'lockup_period': placement_params['lockup_period'],
        'current_price': placement_params['current_price'],
        'risk_free_rate': placement_params['risk_free_rate'],
    }

    # 如果市场数据存在，使用市场数据中的最新价格
    if market_data and 'current_price' in market_data:
        project_params['current_price'] = market_data['current_price']
        print(f"✅ 使用市场数据中的最新价格: {market_data['current_price']:.2f} 元")

    # 4. 构建风险参数
    if market_data:
        # 使用真实市场数据
        risk_params = {
            'volatility': market_data['volatility'],      # 默认60日波动率
            'drift': market_data['drift'],                # 默认60日收益率
            'volatility_30d': market_data.get('volatility_30d'),
            'volatility_60d': market_data.get('volatility_60d'),
            'volatility_120d': market_data.get('volatility_120d'),
            'volatility_180d': market_data.get('volatility_180d'),
            'data_source': 'market_data',
        }
        print(f"✅ 使用真实市场数据:")
        print(f"   波动率: {risk_params['volatility']*100:.2f}% (60日)")
        print(f"   收益率: {risk_params['drift']*100:.2f}% (60日年化)")
    else:
        # 使用定增参数文件中的数据
        risk_params = {
            'volatility': placement_params.get('volatility', 0.35),
            'drift': 0.08,  # 默认假设
            'data_source': 'placement_params',
        }
        print(f"⚠️ 使用参数文件数据:")
        print(f"   波动率: {risk_params['volatility']*100:.2f}%")
        print(f"   收益率: {risk_params['drift']*100:.2f}% (假设)")

    return project_params, risk_params, market_data


def print_config_summary(project_params: Dict, risk_params: Dict, market_data: Dict = None):
    """
    打印配置摘要

    参数:
        project_params: 定增项目参数
        risk_params: 风险参数
        market_data: 市场数据（可选，用于判断发行类型）
    """
    print("\n" + "="*70)
    print("📊 定增分析配置")
    print("="*70)

    print(f"\n📋 项目参数:")
    print(f"   发行价格: {project_params['issue_price']:.2f} 元/股")
    print(f"   当前价格: {project_params['current_price']:.2f} 元/股")
    print(f"   锁定期: {project_params['lockup_period']} 个月")
    print(f"   发行数量: {project_params['issue_shares']:,} 股")
    print(f"   融资金额: {project_params['issue_price'] * project_params['issue_shares'] / 100000000:.2f} 亿元")

    if project_params['current_price'] > project_params['issue_price']:
        current_return = (project_params['current_price'] / project_params['issue_price'] - 1) * 100
        print(f"   当前收益率: {current_return:+.2f}% （浮盈）")
    else:
        current_return = (project_params['current_price'] / project_params['issue_price'] - 1) * 100
        print(f"   当前收益率: {current_return:+.2f}% （浮亏）")

    # 判断发行类型（如果有市场数据）
    if market_data and 'ma_30' in market_data:
        ma30 = market_data['ma_30']
        issue_price = project_params['issue_price']

        print(f"\n📌 发行类型判断:")
        print(f"   MA30: {ma30:.2f} 元")
        print(f"   发行价: {issue_price:.2f} 元")

        if issue_price < ma30:
            safety_margin = (ma30 - issue_price) / ma30 * 100
            print(f"   ✅ 折价发行（有安全边际）")
            print(f"   安全边际: {safety_margin:.2f}%")
        else:
            premium_rate = (issue_price - ma30) / ma30 * 100
            print(f"   ⚠️ 溢价发行（无安全边际）")
            print(f"   溢价幅度: {premium_rate:.2f}%")

    print(f"\n⚠️ 风险参数:")
    print(f"   波动率: {risk_params['volatility']*100:.2f}%")
    print(f"   收益率(漂移率): {risk_params['drift']*100:.2f}%")
    print(f"   数据来源: {risk_params.get('data_source', 'unknown')}")

    if risk_params.get('volatility_30d'):
        print(f"\n📈 波动率详情:")
        print(f"   30日: {risk_params['volatility_30d']*100:.2f}%")
        print(f"   60日: {risk_params['volatility_60d']*100:.2f}%")
        print(f"   120日: {risk_params['volatility_120d']*100:.2f}%")
        print(f"   180日: {risk_params['volatility_180d']*100:.2f}%")

    print("="*70)


# 便捷函数
def quick_load(stock_code: str = '300735.SZ'):
    """
    快速加载配置（一行代码）

    使用示例:
        >>> from utils.config_loader import quick_load
        >>> project_params, risk_params, market_data = quick_load('300735.SZ')
    """
    project_params, risk_params, market_data = load_placement_config(stock_code)
    print_config_summary(project_params, risk_params, market_data)
    return project_params, risk_params, market_data


if __name__ == '__main__':
    # 测试
    try:
        project_params, risk_params, market_data = load_placement_config('300735.SZ')
        print_config_summary(project_params, risk_params, market_data)

        print("\n📌 在其他脚本中使用:")
        print("```python")
        print("from utils.config_loader import load_placement_config")
        print("from utils.analysis_tools import PrivatePlacementRiskAnalyzer")
        print()
        print("# 加载配置")
        print("project_params, risk_params, market_data = load_placement_config('300735.SZ')")
        print()
        print("# 创建分析器")
        print("analyzer = PrivatePlacementRiskAnalyzer(**project_params)")
        print()
        print("# 运行模拟")
        print("sim_results = analyzer.monte_carlo_simulation(")
        print("    n_simulations=10000,")
        print("    volatility=risk_params['volatility'],")
        print("    drift=risk_params['drift'],")
        print("    seed=42")
        print(")")
        print("```")

    except FileNotFoundError as e:
        print(f"❌ {e}")
        print("\n请确保:")
        print("1. 已运行 fetch_gh_data.py 生成定增参数文件")
        print("2. 已运行 07_market_data_analysis.ipynb 生成市场数据文件")
