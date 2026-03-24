#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速测试所有章节模块的导入和基本功能
"""

import sys
import os

# 添加路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MODULE_DIR = os.path.join(SCRIPT_DIR, 'generate_word_report_v2')
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)
sys.path.insert(0, MODULE_DIR)  # 添加模块目录
os.chdir(PROJECT_DIR)

print("="*70)
print("🧪 快速测试所有章节模块")
print("="*70)

# 测试导入
modules = [
    ('module_utils', '工具函数模块'),
    ('chapter01_overview', '第一章：项目概况'),
    ('chapter02_valuation', '第二章：相对估值分析'),
    ('chapter03_dcf', '第三章：DCF估值分析'),
    ('chapter04_sensitivity', '第四章：敏感性分析'),
    ('chapter05_montecarlo', '第五章：蒙特卡洛模拟'),
    ('chapter06_scenario', '第六章：情景分析'),
    ('chapter07_stress', '第七章：压力测试'),
    ('chapter08_var', '第八章：VaR风险度量'),
    ('chapter09_advice', '第九章：风控建议'),
    ('appendix_scenarios', '附件：情景数据表'),
]

success_count = 0
total_count = len(modules)

for module_name, description in modules:
    try:
        module = __import__(f'generate_word_report_v2.{module_name}', fromlist=[module_name])

        # 检查是否有generate_chapter函数
        if hasattr(module, 'generate_chapter'):
            print(f"✅ {module_name:25} - {description}")
            success_count += 1
        else:
            print(f"⚠️  {module_name:25} - {description} (缺少generate_chapter函数)")
    except Exception as e:
        print(f"❌ {module_name:25} - {description}")
        print(f"   错误: {e}")

print("\n" + "="*70)
print(f"导入测试结果: {success_count}/{total_count} 成功")
print("="*70)

# 测试主要函数的签名
print("\n📋 检查主要函数...")

try:
    from generate_word_report_v2.main import generate_report
    print("✅ generate_report 函数可导入")
except Exception as e:
    print(f"❌ generate_report 导入失败: {e}")

try:
    from generate_word_report_v2 import module_utils
    funcs = ['add_title', 'add_paragraph', 'add_table_data', 'add_image']
    for func in funcs:
        if hasattr(module_utils, func):
            print(f"✅ module_utils.{func}")
        else:
            print(f"❌ module_utils.{func} 未找到")
except Exception as e:
    print(f"❌ module_utils 导入失败: {e}")

print("\n" + "="*70)
print("测试完成！")
print("="*70)
