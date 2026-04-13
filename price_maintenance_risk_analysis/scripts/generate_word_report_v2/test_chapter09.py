#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速测试脚本 - 仅测试第九章风控建议与风险提示

用法：
    python test_chapter09.py
"""

import sys
import os
from datetime import datetime
from docx import Document

# 添加项目路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))
sys.path.insert(0, PROJECT_DIR)

# 切换到项目根目录
os.chdir(PROJECT_DIR)

# 导入必要的模块
from utils.config_loader import load_placement_config
from utils.font_manager import get_font_prop

# 导入当前目录的工具函数
import module_utils as utils
import chapter09_advice

def quick_test_chapter09():
    """快速测试第九章生成"""
    print("="*70)
    print(" 第九章快速测试")
    print("="*70)
    print(f"测试时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*70)

    # 加载配置数据
    print("\n 加载配置数据...")
    stock_code = '300735.SZ'
    project_params, risk_params, market_data = load_placement_config(stock_code)

    # 创建Word文档
    document = Document()
    utils.setup_chinese_font(document)

    # 创建context字典
    context = {
        'results': {},
        'project_params': project_params,
        'risk_params': risk_params,
        'market_data': market_data
    }

    # 模拟一些必要的context数据
    # 这里需要模拟之前章节的结果数据

    try:
        # 调用第九章生成函数
        print("\n 生成第九章...")
        chapter09_advice.add_chapter_9(document, context)

        # 保存测试文档
        test_output = os.path.join(PROJECT_DIR, 'reports/outputs', 'test_chapter09.docx')
        document.save(test_output)

        print(f"\n✅ 第九章测试完成！")
        print(f"   测试文档已保存: {test_output}")
        print(f"   请打开文档查看9.3.2.4节的参数构造场景表格")

    except Exception as e:
        print(f"\n❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    quick_test_chapter09()