#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
逐个测试章节模块
"""

import sys
import os

# 添加路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MODULE_DIR = os.path.join(SCRIPT_DIR, 'generate_word_report_v2')
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)
sys.path.insert(0, MODULE_DIR)
os.chdir(PROJECT_DIR)

print("="*70)
print("🧪 逐个测试章节模块")
print("="*70)

# 导入必要的库
from docx import Document
from module_utils import setup_chinese_font

# 准备基础context
from utils.config_loader import load_placement_config
from utils.analysis_tools import PrivatePlacementRiskAnalyzer
from utils.font_manager import get_font_prop

# 加载数据
stock_code = '300735.SZ'
project_params, risk_params, market_data = load_placement_config(stock_code)

font_prop = get_font_prop()
document = Document()
setup_chinese_font(document)

context = {
    'stock_code': stock_code,
    'stock_name': '光弘科技',
    'project_params': project_params,
    'risk_params': risk_params,
    'market_data': market_data,
    'analyzer': PrivatePlacementRiskAnalyzer(
        project_params['issue_price'],
        project_params['current_price'],
        project_params['lockup_period']
    ),
    'document': document,
    'font_prop': font_prop,
    'IMAGES_DIR': '/Users/davy/github/investment_valuation/price_maintenance_risk_analysis/reports/images',
    'DATA_DIR': '/Users/davy/github/investment_valuation/price_maintenance_risk_analysis/data',
    'OUTPUTS_DIR': '/Users/davy/github/investment_valuation/price_maintenance_risk_analysis/reports/outputs',
    'results': {}
}

# 测试各章节
chapters_to_test = [
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

for module_name, description in chapters_to_test:
    print(f"\n📊 测试 {description}...")
    try:
        module = __import__(module_name, fromlist=['generate_chapter'])
        context = module.generate_chapter(context)
        print(f"✅ {description} 测试通过")
    except Exception as e:
        print(f"❌ {description} 测试失败")
        print(f"   错误: {e}")
        import traceback
        traceback.print_exc()
        break

print("\n" + "="*70)
print("测试完成")
print("="*70)
