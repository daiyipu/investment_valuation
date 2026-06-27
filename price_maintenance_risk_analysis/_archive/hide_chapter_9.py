#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
隐藏风险评分章节并调整章节编号
"""

import os

# 读取原文件
file_path = os.path.join(os.path.dirname(__file__), 'scripts/generate_word_report_v2.py')
with open(file_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

# 处理每一行
new_lines = []
skip_chapter_9 = False
in_chapter_9 = False
line_count = len(lines)

for i, line in enumerate(lines):
    line_num = i + 1

    # 检测是否进入第九章
    if "# ==================== 九、综合评估 ====================" in line:
        skip_chapter_9 = True
        in_chapter_9 = True
        new_lines.append(line)
        new_lines.append('    # ==================== 风险评分章节已隐藏 ====================\n')
        new_lines.append('    # 如需启用，请取消下方注释\n')
        new_lines.append('    """\n')
        continue

    # 跳过第九章的内容（直到第十章）
    if skip_chapter_9:
        if line.strip().startswith('add_title') or line.strip().startswith('add_paragraph') or \
           line.strip().startswith('add_table') or line.strip().startswith('add_section'):
            # 注释掉这些行
            new_lines.append(f'    # {line.lstrip()}')
        else:
            new_lines.append(line)

        # 检测是否离开第九章
        if '# ==================== 十、风控建议与风险提示 ====================' in line:
            skip_chapter_9 = False
            in_chapter_9 = False
            new_lines.append('    """\n')  # 结束第九章的注释

        continue

    # 修改章节编号
    if 'add_title(document, \'十、' in line or 'add_title(document, "十、' in line:
        line = line.replace('十、', '九、')
        line = line.replace('10.', '9.')
    elif 'add_title(document, \'10.' in line or 'add_title(document, "10.' in line:
        line = line.replace('10.', '9.')
    elif 'add_title(document, "10.' in line:
        line = line.replace('10.', '9.')
    elif '第十章' in line:
        line = line.replace('第十章', '第九章')
    elif '图表 10.' in line:
        line = line.replace('图表 10.', '图表 9.')

    new_lines.append(line)

# 写回文件
with open(file_path, 'w', encoding='utf-8') as f:
    f.writelines(new_lines)

print("✅ 处理完成：")
print("1. 已隐藏第九章（风险评分体系）")
print("2. 已将第十章改为第九章")
print("3. 已更新图表编号（10.x → 9.x）")
print(f"\n处理了 {line_count} 行代码")
