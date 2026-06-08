#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量定增筛选 + 财务评分 一体化脚本

串联 batch_screener.py（定增决策筛选）和 batch_financial_score.py（财务评分），
输出带评分的最终结果文件。

用法:
  python scripts/batch_screen_and_score.py --input stocks.xlsx [--sheet 0]

流程:
  Step 1: batch_screener.py  → data/batch_screening_result_<date>.xlsx
  Step 2: batch_financial_score.py → data/batch_screening_result_<date>_scored.xlsx

最终输出: data/batch_screening_result_<date>_scored.xlsx
"""

import os
import sys
import time
import subprocess
from datetime import datetime

import openpyxl


# 两个脚本的绝对路径
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
SCREENER_SCRIPT = os.path.join(SCRIPTS_DIR, 'batch_screener.py')

EFAES_SCRIPT = os.path.join(
    os.path.expanduser('~'), 'github', 'EFAES', 'scripts', 'batch_financial_score.py'
)

DATA_DIR = os.path.join(os.path.dirname(SCRIPTS_DIR), 'data')


def run_step(description, cmd):
    """运行一个子步骤，打印分隔线和结果"""
    print()
    print('=' * 60)
    print(f'  {description}')
    print('=' * 60)
    print(f'  命令: {" ".join(cmd)}')
    print()

    t0 = time.time()
    result = subprocess.run(cmd, cwd=DATA_DIR)
    elapsed = time.time() - t0

    if result.returncode != 0:
        print(f'\n❌ {description} 失败 (退出码: {result.returncode}) [耗时 {elapsed:.1f}s]')
        return False

    print(f'\n✅ {description} 完成 [耗时 {elapsed:.1f}s]')
    return True


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description='批量定增筛选 + 财务评分 一体化',
        usage='%(prog)s --input stocks.xlsx [--sheet 0]',
    )
    parser.add_argument('--input', required=True, help='输入Excel文件路径（含"股票代码"和"股票简称"列）')
    parser.add_argument('--sheet', default='0', help='读取Excel的第几个sheet（序号从0开始，默认0）')
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f'❌ 文件不存在: {args.input}')
        sys.exit(1)

    if not os.path.exists(SCREENER_SCRIPT):
        print(f'❌ 找不到 batch_screener.py: {SCREENER_SCRIPT}')
        sys.exit(1)

    if not os.path.exists(EFAES_SCRIPT):
        print(f'❌ 找不到 batch_financial_score.py: {EFAES_SCRIPT}')
        sys.exit(1)

    # 生成带日期+输入文件名的输出文件名（输出到输入文件所在目录）
    date_tag = datetime.now().strftime('%y%m%d')
    input_basename = os.path.splitext(os.path.basename(args.input))[0]
    input_dir = os.path.dirname(os.path.abspath(args.input))
    screening_output = os.path.join(input_dir, f'batch_screening_result_{date_tag}_{input_basename}.xlsx')

    print(f'📁 输入文件: {args.input}')
    print(f'📁 筛选结果: {screening_output}')
    print(f'📁 最终结果: {screening_output.replace(".xlsx", "_scored.xlsx")}')

    t_total = time.time()  # 总计时起点

    # ── Step 1: 定增决策筛选 ──
    step1_cmd = [
        sys.executable, SCREENER_SCRIPT,
        '--input', args.input,
        '--output', screening_output,
        '--sheet', args.sheet,
    ]
    if not run_step('Step 1/2: 定增决策筛选 (batch_screener)', step1_cmd):
        sys.exit(1)

    # 检查筛选结果文件
    if not os.path.exists(screening_output):
        print(f'❌ 筛选结果文件未生成: {screening_output}')
        sys.exit(1)

    if not run_step('Step 2/3: 财务评分 (batch_financial_score)', [
        sys.executable, EFAES_SCRIPT,
        screening_output,
    ]):
        sys.exit(1)

    # 最终输出
    scored_output = screening_output.replace('.xlsx', '_scored.xlsx')
    if not os.path.exists(scored_output):
        print(f'⚠️ 评分结果文件未生成: {scored_output}')
        sys.exit(1)

    # ── Step 3: 追加最终结论列 ──
    # 规则: 定增决策="建议参与" 且 成长能力_趋势="通过" → 最终"通过"
    t_step3 = time.time()
    print()
    print('=' * 60)
    print('  Step 3/3: 生成最终结论')
    print('=' * 60)

    wb = openpyxl.load_workbook(scored_output)
    ws = wb.active

    # 定位关键列
    header_row = [ws.cell(1, col).value for col in range(1, ws.max_column + 1)]
    decision_col = header_row.index('定增决策') + 1
    growth_col = header_row.index('成长能力_趋势') + 1

    # 写入新列
    final_col = ws.max_column + 1
    ws.cell(1, final_col, '最终结论')

    pass_count = 0
    for row_idx in range(2, ws.max_row + 1):
        decision = str(ws.cell(row_idx, decision_col).value or '')
        growth = str(ws.cell(row_idx, growth_col).value or '')

        # 定增决策: 包含"建议参与"即为通过
        decision_pass = '建议参与' in decision
        # 成长能力趋势: 精确匹配"通过"
        growth_pass = growth.strip() == '通过'

        conclusion = '通过' if (decision_pass and growth_pass) else '不通过'
        ws.cell(row_idx, final_col, conclusion)
        if conclusion == '通过':
            pass_count += 1

        stock_name = ws.cell(row_idx, header_row.index('股票简称') + 1).value or ''
        stock_code = ws.cell(row_idx, header_row.index('股票代码') + 1).value or ''
        print(f'  {stock_code} {stock_name}: 定增={"✓" if decision_pass else "✗"} 成长={"✓" if growth_pass else "✗"} → {conclusion}')

    wb.save(scored_output)

    step3_elapsed = time.time() - t_step3
    total = ws.max_row - 1
    print()
    print(f'  最终结论: {pass_count}/{total} 通过 [耗时 {step3_elapsed:.1f}s]')
    print()
    total_elapsed = time.time() - t_total
    print('=' * 60)
    print(f'🎉 全部完成! [总耗时 {total_elapsed:.1f}s]')
    print(f'   筛选结果: {screening_output}')
    print(f'   最终结果: {scored_output}')
    print('=' * 60)


if __name__ == '__main__':
    main()
