#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
东方财富行业板块及成份股数据采集脚本

功能：
  1. 获取东方财富全部行业板块列表（约496个）
  2. 逐个获取每个行业板块的成份股
  3. 存入 valuation.db 数据库

使用：
  python fetch_em_industry_stocks.py              # 全量采集
  python fetch_em_industry_stocks.py --boards-only # 仅采集行业板块列表
  python fetch_em_industry_stocks.py --board BK1027 # 仅采集指定板块的成份股
  python fetch_em_industry_stocks.py --query 600519 # 查询股票所属行业

注意：
  - 需要安装 akshare: pip install akshare
  - 本脚本会自动绕过系统代理，直接连接 push2.eastmoney.com
  - 如果网络不稳定，脚本会自动重试（最多5次）
  - 采集过程中支持断点续采（跳过已有数据的板块）
"""

import argparse
import os
import sys
import time

# ==================== 网络环境修复 ====================

# 清除代理环境变量，避免 Privoxy/ClashX 等本地代理干扰 push2 API
for k in ['http_proxy', 'https_proxy', 'HTTP_PROXY', 'HTTPS_PROXY',
          'all_proxy', 'ALL_PROXY']:
    os.environ.pop(k, None)
os.environ['no_proxy'] = '*'
os.environ['NO_PROXY'] = '*'

# Monkey-patch akshare 的 request_with_retry，添加 trust_env=False
import random

import requests
from requests.adapters import HTTPAdapter

import akshare.utils.request as _req_mod

_original_request_with_retry = _req_mod.request_with_retry


def _patched_request_with_retry(
    url, params=None, timeout=15, max_retries=3,
    base_delay=1.0, random_delay_range=(0.5, 1.5),
):
    """带 trust_env=False 的请求，绕过系统代理。"""
    last_exception = None
    for attempt in range(max_retries):
        try:
            with requests.Session() as session:
                session.trust_env = False
                adapter = HTTPAdapter(pool_connections=1, pool_maxsize=1)
                session.mount("http://", adapter)
                session.mount("https://", adapter)
                response = session.get(url, params=params, timeout=timeout)
                response.raise_for_status()
                return response
        except requests.RequestException as e:
            last_exception = e
            if attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt) + random.uniform(*random_delay_range)
                print(f"    重试 {attempt + 1}/{max_retries}，等待 {delay:.1f}s ...")
                time.sleep(delay)
    raise last_exception


_req_mod.request_with_retry = _patched_request_with_retry

# ==================== 主逻辑 ====================

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.db_manager import ValuationDB

import akshare as ak


def fetch_boards(db):
    """获取全部行业板块列表并存入数据库。"""
    print("=" * 60)
    print("步骤 1: 获取东方财富行业板块列表")
    print("=" * 60)

    for attempt in range(5):
        try:
            df = ak.stock_board_industry_name_em()
            print(f"  获取到 {len(df)} 个行业板块")
            db.save_em_industry_boards(df)
            print(f"  已保存到数据库")
            return df
        except Exception as e:
            wait = (attempt + 1) * 3
            print(f"  第 {attempt + 1} 次获取失败: {type(e).__name__}: {e}")
            if attempt < 4:
                print(f"  等待 {wait}s 后重试...")
                time.sleep(wait)

    print("  ❌ 获取行业板块列表失败，请检查网络连接")
    return None


def fetch_constituent_stocks(db, board_code, board_name):
    """获取指定行业板块的成份股并存入数据库。

    :return: 成份股数量，失败返回 -1
    """
    for attempt in range(3):
        try:
            df = ak.stock_board_industry_cons_em(symbol=board_code)
            if df is not None and len(df) > 0:
                db.save_em_industry_stocks(board_code, df)
                return len(df)
            return 0
        except Exception as e:
            if attempt < 2:
                time.sleep(2 * (attempt + 1))
    return -1


def fetch_all_stocks(db, boards_df=None):
    """获取所有行业板块的成份股。"""
    print()
    print("=" * 60)
    print("步骤 2: 获取各行业板块成份股")
    print("=" * 60)

    # 获取板块列表
    if boards_df is None:
        boards = db.get_em_industry_boards()
        if not boards:
            print("  ❌ 数据库中没有行业板块数据，请先运行 --boards-only")
            return
    else:
        boards = [
            {'board_code': row['板块代码'], 'board_name': row['板块名称']}
            for _, row in boards_df.iterrows()
        ]

    total = len(boards)
    success = 0
    failed = []
    skipped = 0

    # 查找已完成的板块（断点续采）
    existing = db.get_em_industry_boards_count()
    boards_with_stocks = set()
    if existing['boards_with_stocks'] > 0:
        conn = db.get_connection()
        try:
            cur = conn.execute(
                "SELECT board_code FROM em_industry_boards WHERE total_count > 0"
            )
            boards_with_stocks = {row[0] for row in cur.fetchall()}
        finally:
            conn.close()

    print(f"  共 {total} 个板块，已有 {len(boards_with_stocks)} 个板块的成份股数据")

    for i, board in enumerate(boards, 1):
        code = board['board_code']
        name = board['board_name']

        # 跳过已有数据的板块（除非 --force）
        if code in boards_with_stocks:
            skipped += 1
            continue

        count = fetch_constituent_stocks(db, code, name)
        if count >= 0:
            success += 1
            print(f"  [{i}/{total}] {name} ({code}): {count} 只成份股")
        else:
            failed.append(f"{name}({code})")
            print(f"  [{i}/{total}] {name} ({code}): ❌ 失败")

        # 请求间隔，避免触发限频
        time.sleep(0.5)

    print()
    print(f"  采集完成: 成功 {success}, 跳过 {skipped}, 失败 {len(failed)}")
    if failed:
        print(f"  失败板块: {', '.join(failed[:20])}")
        if len(failed) > 20:
            print(f"  ... 以及其他 {len(failed) - 20} 个")


def query_stock(db, stock_code):
    """查询股票所属行业。"""
    industries = db.get_stock_industries(stock_code)
    if industries:
        print(f"\n股票 {stock_code} 所属东方财富行业板块:")
        for ind in industries:
            print(f"  - {ind['board_name']} ({ind['board_code']})")
    else:
        print(f"\n未找到股票 {stock_code} 的行业板块数据")
        print("  请先运行采集脚本：python fetch_em_industry_stocks.py")


def show_stats(db):
    """显示数据库统计信息。"""
    stats = db.get_em_industry_boards_count()
    print("\n数据库统计:")
    print(f"  行业板块数: {stats['board_count']}")
    print(f"  已采集板块: {stats['boards_with_stocks']}")
    print(f"  成份股记录: {stats['stock_count']}")
    print(f"  不重复股票: {stats['unique_stock_count']}")

    # 展示前10个板块
    boards = db.get_em_industry_boards()
    if boards:
        print(f"\n前10个板块:")
        for b in boards[:10]:
            print(f"  {b['board_code']} {b['board_name']} ({b['total_count']} 只)")


def main():
    parser = argparse.ArgumentParser(description='东方财富行业板块及成份股数据采集')
    parser.add_argument('--boards-only', action='store_true',
                        help='仅采集行业板块列表，不获取成份股')
    parser.add_argument('--board', type=str,
                        help='仅采集指定板块的成份股 (如 BK1027)')
    parser.add_argument('--query', type=str,
                        help='查询股票所属行业 (如 600519)')
    parser.add_argument('--force', action='store_true',
                        help='强制重新采集（不跳过已有数据）')
    parser.add_argument('--stats', action='store_true',
                        help='显示数据库统计信息')
    args = parser.parse_args()

    db = ValuationDB()

    if args.stats:
        show_stats(db)
        return

    if args.query:
        query_stock(db, args.query)
        return

    if args.board:
        # 采集单个板块
        conn = db.get_connection()
        try:
            board = conn.execute(
                "SELECT board_code, board_name FROM em_industry_boards WHERE board_code = ?",
                (args.board,)
            ).fetchone()
        finally:
            conn.close()
        if board:
            print(f"采集板块: {board[1]} ({board[0]})")
            count = fetch_constituent_stocks(db, board[0], board[1])
            if count >= 0:
                print(f"  成功获取 {count} 只成份股")
            else:
                print(f"  ❌ 获取失败")
        else:
            print(f"未找到板块 {args.board}，请先运行全量采集")
        return

    # 全量采集
    boards_df = fetch_boards(db)

    if args.boards_only:
        if boards_df is not None:
            show_stats(db)
        return

    if boards_df is not None:
        if args.force:
            # 清空旧数据
            conn = db.get_connection()
            try:
                conn.execute("DELETE FROM em_industry_stocks")
                conn.execute("UPDATE em_industry_boards SET total_count = 0")
                conn.commit()
            finally:
                conn.close()

        fetch_all_stocks(db, boards_df)
    elif db.get_em_industry_boards_count()['board_count'] > 0:
        # 板块列表获取失败，但数据库中有旧数据，继续采集成份股
        print("\n从数据库加载已有板块列表，继续采集成份股...")
        fetch_all_stocks(db)
    else:
        print("\n❌ 无法获取行业板块数据，请检查网络后重试")

    show_stats(db)


if __name__ == '__main__':
    main()
