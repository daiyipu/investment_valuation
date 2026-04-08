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
from datetime import datetime, timedelta
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


def check_data_freshness(market_data, max_trading_days=0):
    """
    检查数据是否过期（基于交易日历）

    参数:
        market_data: 市场数据字典，应包含'analysis_date'
        max_trading_days: 允许的最大落后交易日数（默认0天，不允许落后）

    返回:
        bool: 数据是否新鲜（未过期）
        str: 信息/警告信息
    """
    if 'analysis_date' not in market_data:
        return True, "数据文件中未找到analysis_date字段"

    try:
        analysis_date_str = market_data['analysis_date']
        analysis_date = datetime.strptime(analysis_date_str, '%Y%m%d')
        current_date = datetime.now()

        # 计算交易日差距（排除周末）
        trading_days_diff = 0
        temp_date = analysis_date
        while temp_date < current_date:
            temp_date += timedelta(days=1)
            # 只计算周一到周五（0-4），排除周六周日（5-6）
            if temp_date.weekday() < 5:  # 0=周一, 4=周五
                trading_days_diff += 1

        # 如果当天是周末或节假日，可能没有最新数据，给予宽容
        current_weekday = current_date.weekday()
        if current_weekday >= 5:  # 周末
            # 周末时多给2天宽容期
            trading_days_diff = max(0, trading_days_diff - 2)

        if trading_days_diff > max_trading_days:
            warning_msg = f"""
⚠️  数据过期警告：
   数据文件日期：{analysis_date_str}（{analysis_date.strftime('%Y年%m月%d日')} {['周一','周二','周三','周四','周五','周六','周日'][analysis_date.weekday()]}）
   当前系统日期：{current_date.strftime('%Y%m%d')}（{current_date.strftime('%Y年%m月%d日')} {['周一','周二','周三','周四','周五','周六','周日'][current_weekday]}）
   交易日落后：{trading_days_diff}个交易日
   最大允许：{max_trading_days}个交易日

   建议：请运行以下命令更新数据：
   python scripts/update_indices_data.py

   注意：股票价格实时变动，使用过期数据可能导致分析结果不准确！
            """
            return False, warning_msg.strip()
        else:
            if trading_days_diff == 0:
                return True, "数据最新，无交易日落后"
            else:
                return True, f"数据新鲜，落后{trading_days_diff}个交易日（在{max_trading_days}个交易日允许范围内）"

    except Exception as e:
        return True, f"无法解析数据日期：{e}"


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

    # 检查数据新鲜度
    print("\n 检查数据新鲜度...")
    is_fresh, data_msg = check_data_freshness(market_data)
    print(f" {data_msg}")

    if not is_fresh:
        print("="*70)
        print(" 警告：数据已过期，建议更新后重新生成报告！")
        print("="*70)
        user_input = input("\n是否继续使用过期数据生成报告？(yes/no): ").strip().lower()
        if user_input not in ['yes', 'y', '是']:
            print("报告生成已取消。")
            print("请运行以下命令更新数据后重新生成：")
            print("  python scripts/update_indices_data.py")
            sys.exit(0)

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

    # 生成第九章全部内容（包括9.1综合评估汇总和9.2-9.6其他章节）
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
