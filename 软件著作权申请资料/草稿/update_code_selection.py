#!/usr/bin/env python3
"""
更新代码文件选择，只选择核心系统功能文件，排除报告和模型开发部分。
"""

import json
import sys

def update_selection():
    # 读取当前的选择文件
    with open('代码文件选择.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 定义要选择的文件（核心系统功能）
    selected_files = {
        # 前端入口和路由
        'frontend-vue/src/App.vue': '前端入口组件，体现系统整体架构和导航结构',
        'frontend-vue/src/main.ts': '前端入口文件，Vue应用初始化',
        'frontend-vue/src/router/index.ts': '路由配置，定义系统页面导航结构',

        # 前端核心页面
        'frontend-vue/src/views/Dashboard.vue': '仪表板页面，系统主界面展示',
        'frontend-vue/src/views/History.vue': '历史记录管理页面，查看和对比估值历史',
        'frontend-vue/src/views/ValuationInput.vue': '估值输入页面，核心功能页面，包含估值参数配置和数据处理',
        'frontend-vue/src/views/ValuationResult.vue': '估值结果展示页面，展示多种估值方法结果和风险分析',
        'frontend-vue/src/views/ScenarioAnalysis.vue': '情景分析页面，展示基准、乐观、悲观情景估值对比',
        'frontend-vue/src/views/SensitivityAnalysis.vue': '敏感性分析页面，展示参数敏感性和龙卷风图',
        'frontend-vue/src/views/StressTest.vue': '压力测试页面，展示各类压力测试结果',

        # 前端API服务
        'frontend-vue/src/services/api.ts': '前端API封装，与后端交互的核心接口',

        # 后端API核心文件
        'backend/api/main.py': 'FastAPI应用主入口，定义API路由和CORS配置',
        'backend/api/schemas.py': '数据模型定义，估值请求和响应的数据结构',

        # 后端核心估值服务
        'backend/services/valuation_engine.py': '统一估值引擎，整合多种估值方法的核心算法',
        'backend/services/absolute_valuation.py': '绝对估值服务，DCF现金流折现模型实现',
        'backend/services/relative_valuation.py': '相对估值服务，PE/PS/PB/EV-EBITDA估值倍数计算',
        'backend/services/multi_product_valuation.py': '多产品估值服务，支持复杂业务结构估值',
        'backend/services/scenario_analysis.py': '情景分析服务，多情景估值对比和概率分析',
        'backend/services/sensitivity_analysis.py': '敏感性分析服务，参数敏感性计算和龙卷风图生成',
        'backend/services/stress_test.py': '压力测试服务，蒙特卡洛模拟和风险压力测试',

        # 通用能力文件
        'backend/utils/other_methods.py': '通用工具方法，WACC计算等核心算法'
    }

    # 定义要排除的路径前缀（报告和模型开发部分）
    excluded_prefixes = [
        'enterprise_analysis/',
        'valuation_report/',
        'industry_dcf/',
        'price_models/',
        'price_maintenance_risk_analysis/',
        'ml_training/',
        'agents/',
        'data/',
        'create_ml_features',
        'frontend-vue/public/',
        'frontend-vue/src/components/HelloWorld.vue',
        'scripts/'
    ]

    # 更新选择状态
    selected_count = 0
    total_material_lines = 0

    for file_info in data['files']:
        file_path = file_info['path']

        # 检查是否在选择列表中
        if file_path in selected_files:
            file_info['selected'] = True
            file_info['model_reason'] = selected_files[file_path]
            selected_count += 1
            total_material_lines += file_info['material_line_count']
        # 检查是否在排除列表中
        elif any(file_path.startswith(prefix) for prefix in excluded_prefixes):
            file_info['selected'] = False
            file_info['model_reason'] = '排除：报告和模型开发部分，不属于核心系统功能'
        # 其他文件根据优先级决定
        else:
            # 只保留核心系统功能文件
            if file_info['priority'] <= 30:  # 保留前端核心和后端API服务
                file_info['selected'] = False  # 默认不选，除非在selected_files中
                file_info['model_reason'] = '补充源码文件，如不足60页可考虑补充'
            else:
                file_info['selected'] = False
                file_info['model_reason'] = '排除：低优先级补充文件，非核心系统功能'

    # 更新统计信息
    data['estimated_selected_lines'] = total_material_lines
    data['estimated_selected_pages'] = total_material_lines // data['lines_per_page']

    # 保存更新后的文件
    with open('代码文件选择.json', 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"✅ 代码文件选择已更新")
    print(f"📊 选择了 {selected_count} 个核心文件")
    print(f"📄 预计 {total_material_lines} 行，约 {total_material_lines // data['lines_per_page']} 页")
    print(f"🎯 目标：{data['target_pages']} 页 ({data['target_lines']} 行)")

    # 检查是否达到目标
    if total_material_lines < data['target_lines']:
        shortage = data['target_lines'] - total_material_lines
        print(f"⚠️  还差 {shortage} 行 (约 {shortage // data['lines_per_page']} 页)，需要补充更多文件")
    else:
        print(f"✅ 已达到目标页数要求")

if __name__ == '__main__':
    update_selection()