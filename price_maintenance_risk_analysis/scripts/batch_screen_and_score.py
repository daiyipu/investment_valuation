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
import subprocess
from datetime import datetime


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

    result = subprocess.run(cmd, cwd=DATA_DIR)

    if result.returncode != 0:
        print(f'\n❌ {description} 失败 (退出码: {result.returncode})')
        return False

    print(f'\n✅ {description} 完成')
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

    # 生成带日期的输出文件名
    date_tag = datetime.now().strftime('%y%m%d')
    screening_output = os.path.join(DATA_DIR, f'batch_screening_result_{date_tag}.xlsx')

    print(f'📁 输入文件: {args.input}')
    print(f'📁 筛选结果: {screening_output}')
    print(f'📁 最终结果: {screening_output.replace(".xlsx", "_scored.xlsx")}')

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

    # ── Step 2: 财务评分 ──
    step2_cmd = [
        sys.executable, EFAES_SCRIPT,
        screening_output,
    ]
    if not run_step('Step 2/2: 财务评分 (batch_financial_score)', step2_cmd):
        sys.exit(1)

    # 最终输出
    scored_output = screening_output.replace('.xlsx', '_scored.xlsx')
    print()
    print('=' * 60)
    if os.path.exists(scored_output):
        print(f'🎉 全部完成!')
        print(f'   筛选结果: {screening_output}')
        print(f'   最终结果: {scored_output}')
    else:
        print(f'⚠️ 评分结果文件未生成: {scored_output}')
        sys.exit(1)
    print('=' * 60)


if __name__ == '__main__':
    main()
