#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增风险分析报告生成器 - 统一入口（模块化版本）

功能：
    统一调用各章节模块，生成完整的Word分析报告

使用方法：
    python -m generate_word_report_v2.main
    或
    cd scripts/generate_word_report_v2
    python main.py
"""

import sys
import os
from datetime import datetime
from docx import Document

# 添加项目路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))
sys.path.insert(0, PROJECT_DIR)

# 切换到项目根目录（确保数据文件路径正确）
os.chdir(PROJECT_DIR)

# 导入项目根目录的utils模块
from utils.config_loader import load_placement_config
from utils.analysis_tools import PrivatePlacementRiskAnalyzer
from utils.font_manager import get_font_prop

# 导入本模块的工具函数和章节模块
import module_utils as utils
import chapter01_overview
import chapter02_valuation
import chapter03_dcf
import chapter04_sensitivity
import chapter05_montecarlo
import chapter06_scenario
import chapter07_stress
import chapter08_var
import chapter09_01_evaluation
import chapter09_advice
import appendix_scenarios

# 配置路径
DATA_DIR = os.path.join(PROJECT_DIR, 'data')
REPORTS_DIR = os.path.join(PROJECT_DIR, 'reports')
IMAGES_DIR = os.path.join(REPORTS_DIR, 'images')
OUTPUTS_DIR = os.path.join(REPORTS_DIR, 'outputs')

# 确保目录存在
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(OUTPUTS_DIR, exist_ok=True)

# 获取中文字体
font_prop = get_font_prop()


def generate_report(stock_code='300735.SZ', stock_name='光弘科技'):
    """
    生成完整的定增风险分析报告

    参数:
        stock_code: 股票代码
        stock_name: 股票名称

    返回:
        document: Word文档对象
    """
    print("="*70)
    print(" 定增风险分析报告生成器（模块化版本 V2.0）")
    print("="*70)
    print(f"股票代码: {stock_code}")
    print(f"股票名称: {stock_name}")
    print(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*70)

    # 创建Word文档
    document = Document()
    utils.setup_chinese_font(document)

    # 加载配置数据
    print("\n 加载配置数据...")
    project_params, risk_params, market_data = load_placement_config(stock_code)
    industry_data = _load_industry_data(stock_code)

    # 创建分析器（注意参数顺序：issue_price, issue_shares, lockup_period, current_price, risk_free_rate）
    analyzer = PrivatePlacementRiskAnalyzer(
        project_params['issue_price'],
        project_params['issue_shares'],
        project_params['lockup_period'],
        project_params['current_price'],
        project_params.get('risk_free_rate', 0.03)
    )

    # 共享数据字典（在各章节间传递）
    context = {
        'stock_code': stock_code,
        'stock_name': stock_name,
        'project_params': project_params,
        'risk_params': risk_params,  # 添加risk_params
        'market_data': market_data,
        'industry_data': industry_data,
        'analyzer': analyzer,
        'document': document,
        'font_prop': font_prop,
        'IMAGES_DIR': IMAGES_DIR,
        'DATA_DIR': DATA_DIR,
        'OUTPUTS_DIR': OUTPUTS_DIR,
        # 各章节计算结果（供后续章节使用）
        'results': {}
    }

    # ==================== 依次调用各章节 ====================
    # 第一章：项目概况（包含封面和目录）
    print("\n 生成第一章：项目概况...")
    context = chapter01_overview.generate_chapter(context)

    # 第二章：相对估值分析
    print("\n 生成第二章：相对估值分析...")
    context = chapter02_valuation.generate_chapter(context)

    # 第三章：DCF估值分析
    print("\n 生成第三章：DCF估值分析...")
    context = chapter03_dcf.generate_chapter(context)

    # 第四章：敏感性分析
    print("\n 生成第四章：敏感性分析...")
    context = chapter04_sensitivity.generate_chapter(context)

    # 第五章：蒙特卡洛模拟
    print("\n 生成第五章：蒙特卡洛模拟...")
    context = chapter05_montecarlo.generate_chapter(context)

    # 第六章：情景分析
    print("\n 生成第六章：情景分析...")
    context = chapter06_scenario.generate_chapter(context)

    # 第七章：压力测试
    print("\n 生成第七章：压力测试...")
    context = chapter07_stress.generate_chapter(context)

    # 第八章：VaR风险度量
    print("\n 生成第八章：VaR风险度量...")
    context = chapter08_var.generate_chapter(context)

    # 第九章：风控建议与风险提示
    print("\n 生成第九章：风控建议与风险提示...")

    # 9.1 综合评估汇总
    print("\n 生成第九章第一节：综合评估汇总...")
    context = chapter09_01_evaluation.generate_chapter(context)

    # 9.2-9.9 风控建议与风险提示
    context = chapter09_advice.generate_chapter(context)

    # 附件：情景数据表
    print("\n 生成附件：情景数据表...")
    context = appendix_scenarios.generate_chapter(context)

    print("\n" + "="*70)
    print(" 报告生成完成！")
    print("="*70)

    return document


def _load_industry_data(stock_code):
    """加载行业数据"""
    import json
    industry_data_file = os.path.join(DATA_DIR, f"{stock_code.replace('.', '_')}_industry_data.json")

    if os.path.exists(industry_data_file):
        with open(industry_data_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    else:
        print(f" 未找到行业数据文件: {industry_data_file}")
        return None


def save_report(document, stock_code, stock_name):
    """保存报告到文件"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{stock_name}_{stock_code}_定增风险分析报告_{timestamp}.docx"
    output_path = os.path.join(OUTPUTS_DIR, filename)

    document.save(output_path)
    print(f"\n 报告已保存到: {output_path}")
    return output_path


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='定增风险分析报告生成器（模块化版本）')
    parser.add_argument('--stock', type=str, default='300735.SZ', help='股票代码（默认：300735.SZ）')
    parser.add_argument('--name', type=str, default='光弘科技', help='股票名称（默认：光弘科技）')
    parser.add_argument('--output', type=str, default=None, help='输出文件名（可选）')

    args = parser.parse_args()

    # 生成报告
    doc = generate_report(args.stock, args.name)

    # 保存报告
    if args.output:
        output_path = os.path.join(OUTPUTS_DIR, args.output)
        doc.save(output_path)
        print(f"\n 报告已保存到: {output_path}")
    else:
        save_report(doc, args.stock, args.name)
