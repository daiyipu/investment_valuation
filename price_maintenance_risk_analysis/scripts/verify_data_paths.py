#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
验证所有脚本的数据路径引用是否正确
"""

import os
import sys

# 添加项目路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)

def check_data_paths():
    """检查所有关键文件是否存在"""
    print("=" * 70)
    print("验证数据目录路径")
    print("=" * 70)
    print()

    # 检查data目录
    data_dir = os.path.join(PROJECT_DIR, 'data')
    print(f"✅ 数据目录: {data_dir}")
    print()

    # 检查关键数据文件
    required_files = [
        '300735_SZ_placement_params.json',
        '300735_SZ_market_data.json',
        '300735_SZ_industry_data.json'
    ]

    print("检查关键数据文件:")
    all_exist = True
    for filename in required_files:
        filepath = os.path.join(data_dir, filename)
        exists = os.path.exists(filepath)
        status = "✅" if exists else "❌"
        print(f"  {status} {filename}")
        if not exists:
            all_exist = False

    print()

    # 测试工具函数
    print("测试工具函数:")
    try:
        from utils.config_loader import load_placement_config
        from utils.market_data_loader import load_market_data

        print("  ✅ 导入工具函数成功")

        # 测试加载配置
        try:
            project_params, risk_params, market_data = load_placement_config('300735.SZ')
            print("  ✅ load_placement_config 测试通过")
        except Exception as e:
            print(f"  ❌ load_placement_config 测试失败: {e}")

        # 测试加载数据
        try:
            data = load_market_data('300735.SZ')
            if data:
                print("  ✅ load_market_data 测试通过")
            else:
                print("  ⚠️  load_market_data 返回None（数据文件可能不存在）")
        except Exception as e:
            print(f"  ❌ load_market_data 测试失败: {e}")

    except ImportError as e:
        print(f"  ❌ 导入工具函数失败: {e}")
        all_exist = False

    print()

    # 检查脚本文件
    print("检查脚本文件路径配置:")
    scripts_to_check = [
        'generate_word_report_v2.py',
        'update_market_data.py',
        'update_indices_data.py'
    ]

    for script_name in scripts_to_check:
        script_path = os.path.join(SCRIPT_DIR, script_name)
        if os.path.exists(script_path):
            print(f"  ✅ {script_name}")

            # 读取并检查DATA_DIR定义
            with open(script_path, 'r', encoding='utf-8') as f:
                content = f.read()

            if 'DATA_DIR' in content or 'data_dir' in content:
                if "os.path.join(PROJECT_DIR, 'data')" in content:
                    print(f"     ✅ 使用正确的PROJECT_DIR/data路径")
                elif "os.path.join(os.path.dirname(script_dir), 'data')" in content:
                    print(f"     ✅ 使用正确的相对路径（脚本目录的父目录）")
                elif "'../data'" in content or '"../data"' in content:
                    print(f"     ⚠️  包含../data相对路径（已废弃）")
                else:
                    print(f"     ℹ️  使用其他路径配置")
        else:
            print(f"  ❌ {script_name} 不存在")
            all_exist = False

    print()
    print("=" * 70)
    if all_exist:
        print("✅ 所有检查通过！数据目录配置正确。")
    else:
        print("⚠️  部分检查未通过，请查看上面的详细信息。")
    print("=" * 70)

if __name__ == '__main__':
    check_data_paths()
