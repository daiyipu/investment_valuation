#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
sys.stdout.reconfigure(line_buffering=True)  # 并发线程下强制行缓冲(每行flush)
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
    sub = decision.get('sub_scenarios', {}) or {}

    def _pass(key):
        return '✓' if sub.get(key) else '✗'

    return {
        '股票代码': stock_code,
        '股票简称': stock_name,
        '溢价率下限': f"{pr['min']:+.2f}%" if pr else '-',
        '溢价率上限': f"{pr['max']:+.2f}%" if pr else '-',
        '有效阈值数': decision.get('valid_thresholds', '-'),
        # 子场景明细（放在定增决策前）
        '市场指数': _pass('市场指数'),
        '行业PE': _pass('行业PE'),
        '个股PE': _pass('个股PE'),
        'DCF估值': _pass('DCF估值'),
        '修正PE估值': _pass('修正PE估值'),
        '参数构造': _pass('参数构造'),
        '蒙特卡洛': _pass('蒙特卡洛'),
        '反向推算': _pass('反向推算'),
        '定增决策': decision.get('decision', '分析失败') if decision else '分析失败',
    }


def _add_months(ymd_str, months):
    """日期字符串(YYYYMMDD)加N个月，返回YYYYMMDD。"""
    from datetime import datetime
    dt = datetime.strptime(ymd_str, '%Y%m%d')
    m = dt.month - 1 + months
    y = dt.year + m // 12
    m = m % 12 + 1
    d = min(dt.day, 28)  # 避免无效日（如2月30日）
    return dt.replace(year=y, month=m, day=d).strftime('%Y%m%d')


def _calc_post_issue_return(pro, stock_code, issue_date, months=7):
    """计算报价日N个月后相对报价日的涨跌幅（默认7个月≈解禁后1个月）。

    返回: (报价日价格, N个月后价格, 实际目标日期, 涨跌幅) 或全None
    """
    from datetime import datetime, timedelta
    try:
        target_date = _add_months(issue_date, months)
        # 查询窗口：报价日 ~ 目标日后15天（覆盖非交易日找最近交易日）
        end_fetch = (datetime.strptime(target_date, '%Y%m%d') + timedelta(days=15)).strftime('%Y%m%d')
        df = pro.daily(ts_code=stock_code, start_date=issue_date, end_date=end_fetch)
        if df is None or df.empty:
            return None, None, None, None
        df = df.sort_values('trade_date').reset_index(drop=True)

        # 报价日价格（<=报价日的最近交易日）
        before = df[df['trade_date'] <= issue_date]
        if before.empty:
            issue_price = df.iloc[0]['close']
            issue_actual = df.iloc[0]['trade_date']
        else:
            issue_price = before.iloc[-1]['close']
            issue_actual = before.iloc[-1]['trade_date']

        # 目标日价格（>=目标日的最近交易日）
        after = df[df['trade_date'] >= target_date]
        if after.empty:
            return float(issue_price), None, None, None
        target_price = after.iloc[0]['close']
        target_actual = after.iloc[0]['trade_date']

        if issue_price > 0:
            ret = (target_price - issue_price) / issue_price
            return float(issue_price), float(target_price), target_actual, ret
        return float(issue_price), float(target_price), target_actual, None
    except Exception:
        return None, None, None, None


def run_batch_screening(stock_list, output_path=None):
    """批量筛选主函数。

    Args:
        stock_list: [(stock_code, stock_name), ...] 列表
        output_path: 输出 Excel 路径（默认 data/batch_screening_result.xlsx）
    """
    if output_path is None:
        output_path = os.path.join(PROJECT_DIR, 'data', 'batch_screening_result.xlsx')

    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # 断点续传：读取已完成的股票，跳过它们
    results = []
    done_codes = set()
    if os.path.exists(output_path):
        try:
            existing_df = pd.read_excel(output_path)
            if not existing_df.empty and '股票代码' in existing_df.columns:
                done_codes = set(existing_df['股票代码'].astype(str).str.strip())
                results = existing_df.to_dict('records')
                print(f'📎 检测到已有结果({len(done_codes)}只)，断点续传跳过已完成股票')
        except Exception:
            pass

    # 过滤掉已完成的股票
    pending = [t for t in stock_list if str(t[0]).strip() not in done_codes]
    total_all = len(stock_list)
    total = len(pending)
    skipped_count = total_all - total
    if skipped_count > 0:
        print(f'共 {total_all} 只股票，跳过 {skipped_count} 只已完成，本次处理 {total} 只\n')
    else:
        print(f'共 {total} 只股票\n')

    if total == 0:
        print('✅ 所有股票已完成，无需重新处理')
        return pd.DataFrame(results)

    stock_list = pending  # 只处理未完成的

    # 延迟导入
    from main import generate_report_headless

    # 生成批次ID
    from datetime import datetime
    batch_id = datetime.now().strftime('%Y%m%d_%H%M%S')

    raw_results = []  # 保存原始headless结果用于DB（results已含续传数据）
    t_batch_start = time.time()

    # 初始化Tushare（用于计算报价日后7个月涨跌幅）
    pro = None
    if any(len(item) > 2 and item[2] for item in stock_list):
        try:
            import tushare as ts
            ts_token = os.environ.get('TUSHARE_TOKEN', '')
            if ts_token:
                pro = ts.pro_api(ts_token)
        except Exception:
            pro = None

    # ===== 并发处理（多线程，API调用为主，适合I/O并行）=====
    import threading
    from concurrent.futures import ThreadPoolExecutor, as_completed

    _save_lock = threading.Lock()  # Excel/DB写入锁
    _done_count = [0]  # 已完成数（用list包装以便闭包修改）

    def _process_one(item):
        """处理单只股票（线程内执行），返回结果行"""
        code, name = item[0], item[1]
        issue_date = item[2] if len(item) > 2 else None
        t_stock = time.time()
        headless_result = generate_report_headless(code, name, issue_date=issue_date)
        row = _analyze_one(code, name, lambda c, n: headless_result)
        row['报价日'] = issue_date or ''

        # 计算报价日后7个月涨跌幅
        if issue_date and pro:
            issue_p, target_p, target_date_actual, ret = _calc_post_issue_return(pro, code, issue_date, months=7)
            row['报价日价格'] = f"{issue_p:.2f}" if issue_p else '-'
            row['7个月后价格'] = f"{target_p:.2f}" if target_p else '-'
            row['7个月后涨跌幅'] = f"{ret*100:+.2f}%" if ret is not None else '-'
            time.sleep(0.2)
        else:
            row['报价日价格'] = '-'
            row['7个月后价格'] = '-'
            row['7个月后涨跌幅'] = '-'

        elapsed = time.time() - t_stock
        # 线程安全地保存结果
        with _save_lock:
            results.append(row)
            raw_results.append(headless_result)
            _done_count[0] += 1
            n = _done_count[0]
            # 每3只或最后一只增量保存Excel
            if n % 3 == 0 or n == total:
                try:
                    pd.DataFrame(results).to_excel(output_path, index=False, engine='openpyxl')
                except Exception:
                    pass
            # DB保存
            try:
                from utils.db_manager import ValuationDB
                db = ValuationDB()
                db.save_screening_result(batch_id, code, name, headless_result)
            except Exception:
                pass
        print(f'  ✅ [{n}/{total}] {code} {name} ({elapsed:.1f}s)', flush=True)

    MAX_WORKERS = min(5, total)  # 5并发（MySQL支持完全并发）
    print(f'并发模式：{MAX_WORKERS}线程并行处理\n')
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        list(executor.map(_process_one, stock_list))

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


def _normalize_issue_date(value):
    """将报价日期值规范化为 YYYYMMDD 字符串。

    支持 pandas Timestamp / datetime / 字符串(2020-03-11 / 2020/03/11 / 20200311)。
    返回 None 表示无报价日期。
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    if not s or s.lower() in ('nan', 'nat', 'none'):
        return None
    # pandas Timestamp / datetime 对象
    if hasattr(value, 'strftime'):
        return value.strftime('%Y%m%d')
    # 字符串：统一去分隔符
    for sep in ('-', '/', '.'):
        s = s.replace(sep, '')
    return s if len(s) == 8 and s.isdigit() else None


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
        # 识别报价日期列（支持多种命名）
        date_col = None
        for candidate in ('报价日期', '报价日', '发行日', '发行日期'):
            if candidate in df.columns:
                date_col = candidate
                break
        if date_col:
            print(f'检测到报价日期列: {date_col}')
        for _, row in df.iterrows():
            code = str(row['股票代码']).strip()
            name = str(row['股票简称']).strip()
            issue_date = _normalize_issue_date(row[date_col]) if date_col else None
            stock_list.append((code, name, issue_date))

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
