# -*- coding: utf-8 -*-
"""
定增分析配置加载工具
统一加载定增参数和市场数据，供所有脚本使用
"""

import json
import os
from typing import Dict, Optional, Tuple


def create_default_config(stock_code: str, placement_file: str) -> Dict:
    """
    自动生成默认配置文件

    参数:
        stock_code: 股票代码
        placement_file: 配置文件路径

    返回:
        默认配置字典
    """
    default_config = {
        "financing_amount": 100000000,  # 固定1亿元，用于风险评估
        "lockup_period": 6,
        "pricing_method": "ma20_discount_90",
        "premium_rate": -0.10,
        "risk_free_rate": 0.03,
        "net_assets": 0.0,
        "total_debt": 0.0,
        "net_income": 0.0,
        "revenue_growth": 0.15,
        "operating_margin": 0.15,
        "beta": 1.0,
        "historical_fcf_data": {
            "years": 5,
            "year_range": [2020, 2024],
            "data": []
        },
        "_notes": {
            "financing_amount": "投资金额（元）- 固定1亿元，用于风险评估（与实际投资规模无关）",
            "lockup_period": "锁定期（月）- 默认6个月",
            "pricing_method": "定价方式：ma20_discount_90(MA20九折), ma20_discount_85(MA20八五折), ma20_par(MA20平价), custom_premium(自定义溢价率)",
            "premium_rate": "溢价率（负数为折价，正数为溢价）- 默认-0.10表示九折",
            "_auto_generated": "以下参数自动计算，无需手动填写：",
            "issue_price": "自动计算：MA20 × (1 + premium_rate)",
            "current_price": "自动从API获取最新股价"
        }
    }

    # 确保目录存在
    os.makedirs(os.path.dirname(placement_file), exist_ok=True)

    # 保存默认配置
    with open(placement_file, 'w', encoding='utf-8') as f:
        json.dump(default_config, f, indent=2, ensure_ascii=False)

    print(f"✅ 自动生成默认配置: {placement_file}")
    print(f"⚠️  配置说明：投资金额已固定为1亿元（用于风险评估），无需手动修改")

    return default_config


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
        # 优先使用统一的data目录（price_maintenance_risk_analysis/data）
        # 尝试多个可能的路径，优先级从高到低
        possible_dirs = [
            'data',           # price_maintenance_risk_analysis/data (当在项目根目录时)
            '../data',        # price_maintenance_risk_analysis/data (当在scripts目录时)
            '.',              # 当前目录
            '..'              # 父目录（兼容性）
        ]
        for test_dir in possible_dirs:
            test_file = os.path.join(test_dir, f"{stock_code.replace('.', '_')}_placement_params.json")
            if os.path.exists(test_file):
                data_dir = test_dir
                break

        if data_dir is None:
            data_dir = 'data'  # 默认使用data目录

    # 统一文件名格式
    placement_file = os.path.join(data_dir, f"{stock_code.replace('.', '_')}_placement_params.json")

    # 1. 加载定增参数（如果不存在，自动创建默认配置）
    if not os.path.exists(placement_file):
        placement_params = create_default_config(stock_code, placement_file)
    else:
        with open(placement_file, 'r', encoding='utf-8') as f:
            placement_params = json.load(f)
        print(f"✅ 已加载定增参数: {placement_file}")

    # 2. 尝试加载市场数据（如果启用）
    market_data = None
    if prefer_market_data:
        try:
            # 尝试多种导入方式以适应不同的运行上下文
            try:
                from utils.market_data_loader import load_market_data
            except ImportError:
                from market_data_loader import load_market_data
            market_data = load_market_data(stock_code, data_dir=data_dir)
        except Exception as e:
            print(f"⚠️ 市场数据加载失败: {e}")
            print(f"   将使用默认值")

    # 3. 计算或获取发行价
    if market_data and 'ma_20' in market_data:
        # 根据定价方式和溢价率计算发行价
        pricing_method = placement_params.get('pricing_method', 'ma20_discount_90')
        premium_rate = placement_params.get('premium_rate', -0.10)

        if pricing_method == 'ma20_discount_90':
            premium_rate = -0.10
            price_source = "MA20的九折"
        elif pricing_method == 'ma20_discount_85':
            premium_rate = -0.15
            price_source = "MA20的八五折"
        elif pricing_method == 'ma20_discount_80':
            premium_rate = -0.20
            price_source = "MA20的八折"
        elif pricing_method == 'ma20_par':
            premium_rate = 0.0
            price_source = "MA20平价"
        elif pricing_method == 'custom_premium':
            price_source = f"自定义溢价率({premium_rate*100:.1f}%)"
        else:
            # 默认使用MA20的九折
            premium_rate = -0.10
            price_source = "MA20的九折"

        issue_price = market_data['ma_20'] * (1 + premium_rate)
        print(f"✅ 自动计算发行价: {issue_price:.2f} 元/股 ({price_source}, MA20: {market_data['ma_20']:.2f}元)")
    elif 'issue_price' in placement_params and placement_params['issue_price'] > 0:
        # 兼容旧配置：如果配置文件中有issue_price且大于0，使用配置值
        issue_price = placement_params['issue_price']
        print(f"✅ 使用配置文件中的发行价: {issue_price:.2f} 元/股")
    else:
        # 没有市场数据且配置文件中没有发行价，抛出错误
        raise ValueError("无法获取发行价：请确保已获取市场数据（运行 --force-update）或在配置文件中提供 issue_price")

    # 4. 计算发行数量（固定使用1亿元投资金额）
    # 固定投资金额为1亿元，与实际融资金额无关，仅用于风险评估
    financing_amount = 100000000  # 固定1亿元
    issue_shares = int(financing_amount / issue_price)
    print(f"✅ 使用固定投资金额计算发行数量: {issue_shares:,} 股（投资金额: 1亿元 ÷ 发行价: {issue_price:.2f}元/股）")
    print(f"   注：风险分析基于1亿元投资金额，实际投资规模可根据需要调整")

    # 5. 获取当前价格
    if market_data and 'current_price' in market_data:
        current_price = market_data['current_price']
        print(f"✅ 使用市场数据中的最新价格: {current_price:.2f} 元")
    elif 'current_price' in placement_params and placement_params['current_price'] > 0:
        current_price = placement_params['current_price']
        print(f"✅ 使用配置文件中的当前价格: {current_price:.2f} 元")
    else:
        raise ValueError("无法获取当前价格：请确保已获取市场数据（运行 --force-update）或在配置文件中提供 current_price")

    # 6. 构建项目参数
    project_params = {
        'issue_price': issue_price,
        'issue_shares': issue_shares,
        'lockup_period': placement_params['lockup_period'],
        'current_price': current_price,
        'risk_free_rate': placement_params.get('risk_free_rate', 0.03),
        'financing_amount': 100000000,  # 固定1亿元，用于风险评估
    }

    # 4. 构建风险参数
    if market_data:
        # 使用真实市场数据
        risk_params = {
            'volatility': market_data['volatility'],      # 默认60日波动率
            'drift': market_data['drift'],                # 默认60日收益率
            'volatility_20d': market_data.get('volatility_20d'),
            'volatility_60d': market_data.get('volatility_60d'),
            'volatility_120d': market_data.get('volatility_120d'),
            'volatility_250d': market_data.get('volatility_250d'),
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

    if risk_params.get('volatility_20d'):
        print(f"\n📈 波动率详情:")
        print(f"   月度(20日): {risk_params['volatility_20d']*100:.2f}%")
        print(f"   季度(60日): {risk_params['volatility_60d']*100:.2f}%")
        print(f"   半年(120日): {risk_params['volatility_120d']*100:.2f}%")
        print(f"   年度(250日): {risk_params['volatility_250d']*100:.2f}%")

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
