#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
批量更新notebook和脚本中的windows参数
从 [30, 60, 120, 180, 250] 改为 [20, 60, 120, 250]
"""

import json
import os
import re

def update_notebook_windows(file_path):
    """更新notebook文件中的windows参数"""
    print(f"处理文件: {file_path}")

    with open(file_path, 'r', encoding='utf-8') as f:
        nb = json.load(f)

    modified = False

    for cell in nb.get('cells', []):
        if cell['cell_type'] == 'code':
            source = cell['source']
            if isinstance(source, list):
                source = ''.join(source)

            # 替换windows参数
            patterns = [
                (r"'windows':\s*\[30,\s*60,\s*120,\s*180,\s*250\]", "'windows': [20, 60, 120, 250]"),
                (r"'windows':\s*\[30,\s*60,\s*120,\s*180\]", "'windows': [20, 60, 120, 250]"),
                (r"'windows':\s*\[30,\s*60,\s*120,\s*180,\s*250\].*计算窗口.*30日.*60日.*120日.*半年.*180日.*250日",
                 "'windows': [20, 60, 120, 250]  # 计算窗口：20日(月度)、60日(季度)、120日(半年)、250日(年线)"),
                (r"'windows':\s*\[30,\s*60,\s*120,\s*180\].*计算窗口",
                 "'windows': [20, 60, 120, 250]  # 计算窗口：20日(月度)、60日(季度)、120日(半年)、250日(年线)"),
            ]

            for pattern, replacement in patterns:
                new_source = re.sub(pattern, replacement, source)
                if new_source != source:
                    print(f"  ✓ 替换: {pattern[:50]}...")
                    source = new_source
                    modified = True

            if modified:
                cell['source'] = source

    if modified:
        # 备份原文件
        backup_path = file_path + '.backup_windows'
        with open(backup_path, 'w', encoding='utf-8') as f:
            json.dump(nb, f, indent=1, ensure_ascii=False)
        print(f"  ✅ 已备份到: {backup_path}")

        # 保存修改
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(nb, f, indent=1, ensure_ascii=False)
        print(f"  ✅ 已更新: {file_path}")
    else:
        print(f"  ⚠️ 无需修改")

    return modified

def main():
    """主函数"""
    base_dir = '../'

    # 需要更新的notebook文件
    notebook_files = [
        'notebooks_new/10_market_data_analysis.ipynb',
        'notebooks/10_market_data_analysis.ipynb',
        'notebooks/07_market_data_analysis.ipynb',
    ]

    print("="*70)
    print("更新windows参数")
    print("="*70)
    print(f"从: [30, 60, 120, 180, 250]")
    print(f"到: [20, 60, 120, 250]")
    print("="*70)
    print()

    updated_count = 0

    for nb_file in notebook_files:
        file_path = os.path.join(base_dir, nb_file)
        if os.path.exists(file_path):
            try:
                if update_notebook_windows(file_path):
                    updated_count += 1
                print()
            except Exception as e:
                print(f"  ❌ 失败: {e}")
                print()
        else:
            print(f"⚠️ 文件不存在: {file_path}")
            print()

    print("="*70)
    print(f"✅ 完成! 更新了 {updated_count} 个文件")
    print("="*70)

    print("\n📌 更新说明:")
    print("  20日 - 月度指标")
    print("  60日 - 季度指标")
    print("  120日 - 半年指标")
    print("  250日 - 年度指标")

if __name__ == '__main__':
    main()
