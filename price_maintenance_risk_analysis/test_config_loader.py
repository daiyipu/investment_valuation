#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
配置加载器测试脚本
"""
import sys
from utils.config_loader import load_placement_config

print("🧪 测试配置加载器...\n")

try:
    # 加载配置
    project_params, risk_params = load_placement_config('300735.SZ')

    print('\n' + '='*70)
    print('✅ 配置加载成功！')
    print('='*70)

    print(f'\n📋 项目参数:')
    print(f'  发行价格: {project_params["issue_price"]:.2f} 元')
    print(f'  当前价格: {project_params["current_price"]:.2f} 元')
    print(f'  锁定期: {project_params["lockup_period"]} 个月')

    print(f'\n⚠️ 风险参数:')
    print(f'  波动率: {risk_params["volatility"]*100:.2f}%')
    print(f'  收益率: {risk_params["drift"]*100:.2f}%')
    print(f'  数据来源: {risk_params["data_source"]}')

    if risk_params.get('volatility_30d'):
        print(f'\n📊 多窗口波动率:')
        for window in ['30d', '60d', '120d', '180d']:
            vol = risk_params.get(f'volatility_{window}')
            if vol:
                print(f'  {window}: {vol*100:.2f}%')

    print('\n' + '='*70)
    print('📌 在脚本中调用示例:')
    print('='*70)
    print('```python')
    print('from utils.config_loader import load_placement_config')
    print('from utils.analysis_tools import PrivatePlacementRiskAnalyzer')
    print('')
    print('# 加载配置（自动使用真实数据）')
    print('project_params, risk_params = load_placement_config("300735.SZ")')
    print('')
    print('# 创建分析器')
    print('analyzer = PrivatePlacementRiskAnalyzer(**project_params)')
    print('')
    print('# 使用真实参数运行模拟')
    print('sim_results = analyzer.monte_carlo_simulation(')
    print('    n_simulations=10000,')
    print('    volatility=risk_params["volatility"],  # 真实波动率')
    print('    drift=risk_params["drift"],             # 真实收益率')
    print('    seed=42')
    print(')')
    print('```')
    print('='*70)

except Exception as e:
    print(f'❌ 配置加载失败: {e}')
    import traceback
    traceback.print_exc()
