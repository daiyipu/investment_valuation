#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速测试和验证脚本集

提供多种快速验证方式，避免每次生成完整报告：
"""

import sys
import os

# 添加项目路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))
sys.path.insert(0, PROJECT_DIR)
os.chdir(PROJECT_DIR)

def show_usage():
    """显示使用说明"""
    print("="*70)
    print(" 快速测试和验证工具集")
    print("="*70)
    print()
    print("📋 使用方法：")
    print("  1. 验证数据逻辑（最快）：")
    print("     python verify_9_3_2_4.py")
    print()
    print("  2. 测试单个章节（快速）：")
    print("     python test_chapter09.py")
    print()
    print("  3. 生成完整报告（完整验证）：")
    print("     python main.py")
    print()
    print("🔧 优化建议：")
    print("  • 修改代码后先用验证脚本检查逻辑")
    print("  • 确认逻辑正确后再生成完整报告")
    print("  • 使用章节测试脚本验证特定章节")
    print()
    print("="*70)

def quick_verify_9_3_2_4():
    """快速验证9.3.2.4节"""
    print("\n🔍 验证9.3.2.4节参数构造场景...")

    # 导入验证脚本
    import importlib
    verify_module = importlib.import_module('verify_9_3_2_4')
    verify_module.verify_parameter_scenarios()

    print("\n✅ 验证完成！")

def quick_test_chapter09():
    """快速测试第九章"""
    print("\n📄 测试第九章生成...")

    # 导入测试脚本
    import importlib
    test_module = importlib.import_module('test_chapter09')
    test_module.quick_test_chapter09()

    print("\n✅ 测试完成！")

if __name__ == '__main__':
    if len(sys.argv) > 1:
        command = sys.argv[1]

        if command == 'verify':
            quick_verify_9_3_2_4()
        elif command == 'test':
            quick_test_chapter09()
        else:
            print(f"❌ 未知命令: {command}")
            show_usage()
    else:
        show_usage()