#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模块化重构测试脚本

测试各章节模块的语法和导入
"""

import sys
import os
import time

# 添加项目路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)

# 配置路径
MODULE_DIR = os.path.join(SCRIPT_DIR, 'generate_word_report_v2')


def test_imports():
    """测试所有模块是否可以正确导入"""
    print("="*70)
    print("🧪 测试1：模块导入测试")
    print("="*70)

    modules_to_test = [
        ('utils', '工具函数模块'),
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

    results = []

    for module_name, description in modules_to_test:
        try:
            module = __import__(f'generate_word_report_v2.{module_name}', fromlist=[module_name])
            results.append((module_name, description, '✅ 成功', None))
            print(f"✅ {module_name:25} - {description}")
        except Exception as e:
            results.append((module_name, description, '❌ 失败', str(e)))
            print(f"❌ {module_name:25} - {description}")
            print(f"   错误: {e}")

    # 统计
    success_count = sum(1 for r in results if r[2] == '✅ 成功')
    total_count = len(results)

    print("\n" + "="*70)
    print(f"导入测试结果: {success_count}/{total_count} 成功")
    print("="*70)

    return results


def test_function_interfaces():
    """测试各章节模块的函数接口"""
    print("\n" + "="*70)
    print("🧪 测试2：函数接口测试")
    print("="*70)

    modules_to_test = [
        ('utils', ['generate_break_even_chart', 'add_title', 'add_paragraph']),
        ('chapter01_overview', ['generate_chapter']),
        ('chapter02_valuation', ['generate_chapter']),
        ('chapter03_dcf', ['generate_chapter']),
        ('chapter04_sensitivity', ['generate_chapter']),
        ('chapter05_montecarlo', ['generate_chapter']),
        ('chapter06_scenario', ['generate_chapter']),
        ('chapter07_stress', ['generate_chapter']),
        ('chapter08_var', ['generate_chapter']),
        ('chapter09_advice', ['generate_chapter']),
        ('appendix_scenarios', ['generate_chapter']),
    ]

    results = []

    for module_name, expected_functions in modules_to_test:
        try:
            module = __import__(f'generate_word_report_v2.{module_name}', fromlist=[module_name])

            for func_name in expected_functions:
                if hasattr(module, func_name):
                    print(f"✅ {module_name:25}.{func_name}")
                else:
                    print(f"❌ {module_name:25}.{func_name} - 未找到")
                    results.append((module_name, func_name, '❌ 未找到'))
        except Exception as e:
            print(f"❌ {module_name:25} - 导入失败: {e}")

    print("\n" + "="*70)
    print("函数接口测试完成")
    print("="*70)


def test_syntax():
    """测试所有Python文件的语法"""
    print("\n" + "="*70)
    print("🧪 测试3：语法检查")
    print("="*70)

    import py_compile

    # 获取所有Python文件
    python_files = []
    for root, dirs, files in os.walk(MODULE_DIR):
        for file in files:
            if file.endswith('.py'):
                python_files.append(os.path.join(root, file))

    results = []

    for file_path in python_files:
        try:
            py_compile.compile(file_path, doraise=True)
            rel_path = os.path.relpath(file_path, MODULE_DIR)
            print(f"✅ {rel_path}")
            results.append((rel_path, '✅ 成功'))
        except py_compile.PyCompileError as e:
            rel_path = os.path.relpath(file_path, MODULE_DIR)
            print(f"❌ {rel_path}")
            print(f"   错误: {e}")
            results.append((rel_path, f'❌ 失败: {e}'))

    success_count = sum(1 for r in results if '✅' in r[1])
    total_count = len(results)

    print("\n" + "="*70)
    print(f"语法检查结果: {success_count}/{total_count} 成功")
    print("="*70)

    return results


def test_dependencies():
    """测试依赖库是否已安装"""
    print("\n" + "="*70)
    print("🧪 测试4：依赖库检查")
    print("="*70)

    dependencies = [
        ('docx', 'python-docx', 'Word文档生成'),
        ('matplotlib', 'matplotlib', '图表生成'),
        ('numpy', 'numpy', '数值计算'),
        ('pandas', 'pandas', '数据处理'),
        ('scipy', 'scipy', '科学计算'),
        ('tushare', 'tushare', '数据源'),
        ('statsmodels', 'statsmodels', 'ARIMA模型'),
        ('arch', 'arch', 'GARCH模型'),
    ]

    results = []

    for module_name, package_name, description in dependencies:
        try:
            module = __import__(module_name)
            version = getattr(module, '__version__', '未知')
            print(f"✅ {module_name:15} ({package_name:20}) - {description:20} - v{version}")
            results.append((module_name, '✅ 已安装'))
        except ImportError:
            print(f"❌ {module_name:15} ({package_name:20}) - {description:20} - 未安装")
            results.append((module_name, '❌ 未安装'))

    # 统计
    success_count = sum(1 for r in results if '✅' in r[1])
    total_count = len(results)

    print("\n" + "="*70)
    print(f"依赖检查结果: {success_count}/{total_count} 已安装")
    print("="*70)

    if success_count < total_count:
        print("\n⚠️ 缺失依赖库安装命令：")
        print("   pip install statsmodels arch")

    return results


def test_module_size():
    """统计各模块的代码行数"""
    print("\n" + "="*70)
    print("🧪 测试5：代码统计")
    print("="*70)

    # 获取所有Python文件
    python_files = []
    for root, dirs, files in os.walk(MODULE_DIR):
        for file in files:
            if file.endswith('.py') and not file.startswith('__pycache__'):
                python_files.append(os.path.join(root, file))

    total_lines = 0
    file_stats = []

    for file_path in sorted(python_files):
        with open(file_path, 'r', encoding='utf-8') as f:
            line_count = sum(1 for _ in f)
        total_lines += line_count

        rel_path = os.path.relpath(file_path, PROJECT_DIR)
        file_stats.append((rel_path, line_count))
        print(f"{line_count:5d} 行  {rel_path}")

    print("\n" + "-"*70)
    print(f"总计: {total_lines} 行")
    print("="*70)

    return file_stats


def main():
    """运行所有测试"""
    print("\n")
    print("╔" + "="*68 + "╗")
    print("║" + " "*20 + "🧪 模块化重构测试报告" + " "*20 + "║")
    print("╚" + "="*68 + "╝")
    print()

    start_time = time.time()

    # 运行测试
    import_results = test_imports()
    interface_results = test_function_interfaces()
    syntax_results = test_syntax()
    dep_results = test_dependencies()
    size_results = test_module_size()

    end_time = time.time()
    elapsed = end_time - start_time

    # 总结
    print("\n")
    print("╔" + "="*68 + "╗")
    print("║" + " "*25 + "📊 测试总结" + " "*25 + "║")
    print("╚" + "="*68 + "╝")
    print()

    import_success = sum(1 for r in import_results if '✅' in r[2])
    syntax_success = sum(1 for r in syntax_results if '✅' in r[1])
    dep_success = sum(1 for r in dep_results if '✅' in r[1])

    print(f"测试耗时: {elapsed:.2f} 秒")
    print(f"模块导入: {import_success}/{len(import_results)} 成功")
    print(f"语法检查: {syntax_success}/{len(syntax_results)} 成功")
    print(f"依赖安装: {dep_success}/{len(dep_results)} 已安装")

    if import_success == len(import_results) and syntax_success == len(syntax_results):
        print("\n✅ 所有核心测试通过！")
    else:
        print("\n⚠️ 部分测试失败，请检查上述错误信息")

    print("\n" + "="*70)
    print("测试完成")
    print("="*70)


if __name__ == '__main__':
    main()
