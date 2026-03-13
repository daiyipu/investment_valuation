#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
更新市场指数数据文件，添加20日和250日指标
"""

import json
import os

def update_indices_data():
    """
    更新market_indices_scenario_data.json，添加20日和250日指标
    使用30日作为20日的近似值，180日作为250日的近似值
    """
    # 获取脚本所在目录的绝对路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # data目录在项目根目录下（scripts的父目录）
    data_dir = os.path.join(os.path.dirname(script_dir), 'data')
    data_file = os.path.join(data_dir, 'market_indices_scenario_data.json')

    if not os.path.exists(data_file):
        print(f"❌ 文件不存在: {data_file}")
        return

    # 读取现有数据
    with open(data_file, 'r', encoding='utf-8') as f:
        indices_data = json.load(f)

    print(f"✅ 已加载 {len(indices_data)} 个指数数据")

    # 更新每个指数的数据
    for index_name, data in indices_data.items():
        # 添加/更新20日指标（使用30日作为近似值）
        if 'volatility_30d' in data and 'volatility_20d' not in data:
            data['volatility_20d'] = data['volatility_30d']
        if 'return_30d' in data and 'return_20d' not in data:
            data['return_20d'] = data['return_30d']
        if 'ma_30' in data and 'ma_20' not in data:
            data['ma_20'] = data['ma_30']

        # 添加/更新250日指标（使用180日作为近似值）
        if 'volatility_180d' in data and 'volatility_250d' not in data:
            data['volatility_250d'] = data['volatility_180d']
        if 'return_180d' in data and 'return_250d' not in data:
            data['return_250d'] = data['return_180d']
        if 'ma_180' in data and 'ma_250' not in data:
            data['ma_250'] = data['ma_180']

        # 添加win_rate_250d（使用60日胜率作为近似值）
        if 'win_rate_60d' in data and 'win_rate_250d' not in data:
            data['win_rate_250d'] = data['win_rate_60d']

        print(f"  {index_name}: volatility_20d={data.get('volatility_20d', 0):.4f}, volatility_250d={data.get('volatility_250d', 0):.4f}")

    # 保存更新后的数据
    with open(data_file, 'w', encoding='utf-8') as f:
        json.dump(indices_data, f, indent=2, ensure_ascii=False)

    print(f"\n✅ 已更新指数数据文件: {data_file}")
    print("\n📊 更新后的数据结构包含:")
    print("   - volatility_20d, volatility_60d, volatility_120d, volatility_250d")
    print("   - return_20d, return_60d, return_120d, return_250d")
    print("   - ma_20, ma_60, ma_120, ma_250")
    print("   - win_rate_60d, win_rate_250d")

if __name__ == '__main__':
    update_indices_data()
