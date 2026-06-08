#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量定增决策筛选工具

支持两种用法：
  1) 直接指定股票（类似 gen_report.sh）:
     python scripts/batch_screener.py 300604.SZ 长川科技
     python scripts/batch_screener.py 300604.SZ 长川科技 002001.SZ 华兰生物

  2) 从 Excel 批量导入:
     python scripts/batch_screener.py --input stocks.xlsx [--output result.xlsx]

  也可混合使用:
     python scripts/batch_screener.py 300604.SZ 长川科技 --input stocks.xlsx
"""

import argparse
import os
import sys
import time

import pandas as pd

# 添加路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
GENERATE_DIR = os.path.join(SCRIPT_DIR, 'generate_word_report_v2')
sys.path.insert(0, PROJECT_DIR)
sys.path.insert(0, GENERATE_DIR)
os.chdir(PROJECT_DIR)


def _analyze_one(stock_code, stock_name, headless_fn):
    """分析单只股票，返回结果 dict。"""
    result = headless_fn(stock_code, stock_name)
    decision = result.get('decision_conclusion') or {}
    pr = decision.get('premium_range')
    error = result.get('error')

    return {
        '股票代码': stock_code,
        '股票简称': stock_name,
        '溢价率下限': f"{pr['min']:+.2f}%" if pr else '-',
        '溢价率上限': f"{pr['max']:+.2f}%" if pr else '-',
        '有效阈值数': decision.get('valid_thresholds', '-'),
        '①历史数据场景': ('✓ 通过' if decision.get('step1', {}).get('pass') else '✗ 未通过') if decision else '-',
        '②预期估值场景': ('✓ 通过' if decision.get('step2', {}).get('pass') else '✗ 未通过') if decision else '-',
        '③其他场景': ('✓ 通过' if decision.get('step3', {}).get('pass') else '✗ 未通过') if decision else '-',
        '定增决策': decision.get('decision', '分析失败') if decision else '分析失败',
        '决策详情': decision.get('summary', error or '无结果') if decision else (error or '无结果'),
    }


def run_batch_screening(stock_list, output_path=None):
    """批量筛选主函数。

    Args:
        stock_list: [(stock_code, stock_name), ...] 列表
        output_path: 输出 Excel 路径（默认 data/batch_screening_result.xlsx）
    """
    if output_path is None:
        output_path = os.path.join(PROJECT_DIR, 'data', 'batch_screening_result.xlsx')

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    total = len(stock_list)
    print(f'共 {total} 只股票\n')

    # 延迟导入
    from main import generate_report_headless

    # 生成批次ID
    from datetime import datetime
    batch_id = datetime.now().strftime('%Y%m%d_%H%M%S')

    results = []
    raw_results = []  # 保存原始headless结果用于DB
    t_batch_start = time.time()
    for idx, (code, name) in enumerate(stock_list, 1):
        t_stock = time.time()
        print(f'[{idx}/{total}] {code} {name}')
        headless_result = generate_report_headless(code, name)
        stock_elapsed = time.time() - t_stock
        raw_results.append(headless_result)
        results.append(_analyze_one(code, name, lambda c, n: headless_result))
        print(f'  ⏱ {code} 耗时 {stock_elapsed:.1f}s')

        # 保存到DB
        try:
            from utils.db_manager import ValuationDB
            db = ValuationDB()
            db.save_screening_result(batch_id, code, name, headless_result)
        except Exception:
            pass

    batch_elapsed = time.time() - t_batch_start

    # 写入 Excel
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False, engine='openpyxl')
    print(f'\n结果已写入: {output_path}')

    # 打印汇总
    pass_count = sum(1 for r in results if r['定增决策'] == '建议参与本次定向增发')
    fail_count = total - pass_count
    avg_time = batch_elapsed / total if total > 0 else 0
    print(f'汇总: 通过 {pass_count} / 不通过 {fail_count} / 共 {total} [总耗时 {batch_elapsed:.1f}s, 平均 {avg_time:.1f}s/只]')

    return df


def _parse_stock_args(args):
    """从命令行位置参数中解析 股票代码/名称 对。"""
    stock_list = []
    i = 0
    while i < len(args) - 1:
        code = args[i]
        name = args[i + 1]
        if code.startswith('--'):
            break
        stock_list.append((code.strip(), name.strip()))
        i += 2
    return stock_list


def main():
    parser = argparse.ArgumentParser(
        description='批量定增决策筛选',
        usage='%(prog)s [股票代码 股票简称 ...] [--input stocks.xlsx] [--output result.xlsx]',
    )
    parser.add_argument('--input', default=None, help='输入Excel文件路径（含"股票代码"和"股票简称"列）')
    parser.add_argument('--output', default=None, help='输出Excel文件路径')
    parser.add_argument('--sheet', default=0, help='读取Excel的第几个sheet（序号从0开始，默认0）')
    args, remaining = parser.parse_known_args()

    stock_list = []

    # 1) 从位置参数解析（格式: 300604.SZ 长川科技 002001.SZ 华兰生物）
    if remaining:
        stock_list.extend(_parse_stock_args(remaining))

    # 2) 从 Excel 读取
    if args.input:
        print(f'读取输入文件: {args.input}')
        sheet = int(args.sheet)
        df = pd.read_excel(args.input, sheet_name=sheet)
        if '股票代码' not in df.columns or '股票简称' not in df.columns:
            print(f'错误: Excel 需包含"股票代码"和"股票简称"两列')
            print(f'当前列: {list(df.columns)}')
            sys.exit(1)
        for _, row in df.iterrows():
            stock_list.append((str(row['股票代码']).strip(), str(row['股票简称']).strip()))

    if not stock_list:
        print('错误: 请指定股票（位置参数或 --input Excel）')
        print()
        print('用法:')
        print('  python scripts/batch_screener.py 300604.SZ 长川科技')
        print('  python scripts/batch_screener.py 300604.SZ 长川科技 002001.SZ 华兰生物')
        print('  python scripts/batch_screener.py --input stocks.xlsx')
        sys.exit(1)

    run_batch_screening(stock_list, args.output)


if __name__ == '__main__':
    main()
