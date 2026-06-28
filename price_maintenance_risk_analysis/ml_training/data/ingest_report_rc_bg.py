#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""report_rc 后台慢任务 — 断点续接。

report_rc API 限频 1次/小时,~4000 交易日需 ~167 天。
本脚本每次拉取 1 个 report_date 截面,存入 analyst_report_daily 表,
然后从表中计算 PE 特征。支持断点续接(从上次中断处继续)。

用法:
  # 拉取下一个可用的 report_date(自动跳过已拉取的)
  python ingest_report_rc_bg.py fetch [--count N]   # 默认拉1个,count=N拉N个(受限于限频)

  # 从已存数据计算 PE 特征
  python ingest_report_rc_bg.py compute [--limit N]

  # 查看进度
  python ingest_report_rc_bg.py status
"""
import argparse
import json
import os
import sys
import time
from collections import defaultdict

import numpy as np
import pandas as pd
import pymysql

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(PKG, 'ml_training'))
sys.path.insert(0, PKG)
from utils.db_manager import ValuationDB
from tushare_token import resolve_tushare_token
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
import tushare as ts

_CFG = ValuationDB.MYSQL_CONFIG
CHECKPOINT_FILE = os.path.join(os.path.dirname(__file__), '.report_rc_checkpoint.json')

# ── 评级映射 ──
_RATING_MAP = {'买入': 5, '强烈推荐': 5, '推荐': 5, '增持': 4, '谨慎推荐': 4,
               '中性': 3, '持有': 3, '减持': 2, '卖出': 1, '回避': 1}


def _ensure_tables(conn):
    cur = conn.cursor()
    # 原始数据表
    cur.execute("""CREATE TABLE IF NOT EXISTS analyst_report_daily (
        report_date CHAR(8) NOT NULL,
        ts_code VARCHAR(16) NOT NULL,
        name VARCHAR(64),
        org_name VARCHAR(128),
        author_name VARCHAR(128),
        quarter VARCHAR(16),
        eps DOUBLE,
        np DOUBLE COMMENT '预测净利润(万元)',
        rating VARCHAR(16),
        rating_num INT,
        max_price DOUBLE,
        min_price DOUBLE,
        pe DOUBLE,
        roe DOUBLE,
        UNIQUE KEY uk (report_date, ts_code, org_name(64))
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
    # 进度表
    cur.execute("""CREATE TABLE IF NOT EXISTS analyst_report_progress (
        report_date CHAR(8) PRIMARY KEY,
        rows_fetched INT,
        fetched_at DATETIME DEFAULT CURRENT_TIMESTAMP
    ) ENGINE=InnoDB""")
    conn.commit()
    cur.close()


def _get_fetched_dates(conn):
    """已拉取的 report_date 集合。"""
    cur = conn.cursor()
    cur.execute("SELECT report_date FROM analyst_report_progress ORDER BY report_date")
    dates = {r[0] for r in cur.fetchall()}
    cur.close()
    return dates


def _get_trade_dates(conn, start='20100101', end=None):
    """获取交易日历中所有交易日。"""
    pro = ts.pro_api()
    if end is None:
        end = time.strftime('%Y%m%d')
    cal = pro.trade_cal(exchange='SSE', start_date=start, end_date=end, is_open='1')
    if cal is None or len(cal) == 0:
        return []
    return sorted(cal['cal_date'].tolist())


def fetch_next(conn, count=1):
    """拉取下一个未拉取的 report_date 截面。"""
    _ensure_tables(conn)
    fetched = _get_fetched_dates(conn)
    all_dates = _get_trade_dates(conn)
    pending = [d for d in all_dates if d not in fetched]
    if not pending:
        print('✅ 所有交易日已拉取完毕!')
        return 0

    print(f'进度: {len(fetched)}/{len(all_dates)} 已拉取, {len(pending)} 待拉取')
    print(f'下一个: {pending[0]}')

    pro = ts.pro_api()
    cur = conn.cursor()
    total_rows = 0
    for i, rd in enumerate(pending[:count]):
        print(f'  拉取 report_date={rd}...', end='', flush=True)
        try:
            df = pro.report_rc(report_date=rd)
        except Exception as e:
            if '频率超限' in str(e):
                print(f' ⏳ 限频,请等待后重试')
                break
            print(f' ❌ {e}')
            break

        if df is None or len(df) == 0:
            # 空日(周末/假日/无研报),记录为已处理
            cur.execute("INSERT IGNORE INTO analyst_report_progress (report_date, rows_fetched) VALUES (%s, 0)", (rd,))
            conn.commit()
            print(f' 空(无研报)')
            continue

        # 写入 analyst_report_daily
        rows = []
        for _, r in df.iterrows():
            rating = str(r.get('rating', ''))
            rows.append((
                rd,
                str(r.get('ts_code', '')),
                str(r.get('name', ''))[:64],
                str(r.get('org_name', ''))[:128],
                str(r.get('author_name', ''))[:128],
                str(r.get('quarter', '')),
                _sv(r.get('eps')),
                _sv(r.get('np')),
                rating,
                _RATING_MAP.get(rating),
                _sv(r.get('max_price')),
                _sv(r.get('min_price')),
                _sv(r.get('pe')),
                _sv(r.get('roe')),
            ))
        if rows:
            cur.executemany(
                "INSERT IGNORE INTO analyst_report_daily "
                "(report_date, ts_code, name, org_name, author_name, quarter, "
                "eps, np, rating, rating_num, max_price, min_price, pe, roe) "
                "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                rows)
        # 记录进度
        cur.execute(
            "INSERT INTO analyst_report_progress (report_date, rows_fetched) VALUES (%s, %s) "
            "ON DUPLICATE KEY UPDATE rows_fetched=VALUES(rows_fetched), fetched_at=CURRENT_TIMESTAMP",
            (rd, len(rows)))
        conn.commit()
        total_rows += len(rows)
        print(f' ✅ {len(rows)} 条')

        if i < count - 1:
            print(f'  ⏳ 限频等待(1次/小时)...如需拉多个请配合 cron 调度')
            break  # 限频,一次只拉一个

    cur.close()
    print(f'\n本次拉取 {total_rows} 条')
    return total_rows


def compute_features(conn, limit=0):
    """从 analyst_report_daily 计算 PE 特征。"""
    _ensure_tables(conn)
    # 加载样本
    cur = conn.cursor()
    cur.execute("SELECT stock_code, issue_date, issue_date_price FROM placement_evaluation "
                "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8")
    cols = [d[0] for d in cur.description]
    samp = pd.DataFrame(cur.fetchall(), columns=cols)
    cur.close()
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]

    # 确保 PE 列存在
    from ml_training.data.fetch_factors import ensure_columns, batch_update, COLS
    ensure_columns(conn, 'report_rc')

    # 预加载所有研报数据(按 stock 分组)
    print('  加载 analyst_report_daily...')
    cur = conn.cursor()
    cur.execute("SELECT * FROM analyst_report_daily ORDER BY ts_code, report_date")
    acols = [d[0] for d in cur.description]
    all_reports = pd.DataFrame(cur.fetchall(), columns=acols)
    cur.close()
    print(f'  共 {len(all_reports)} 条研报')

    report_map = {s: g for s, g in all_reports.groupby('ts_code')}
    rows = []
    for i, stock in enumerate(stocks):
        df = report_map.get(stock)
        if df is None or len(df) == 0:
            continue
        df['eps'] = pd.to_numeric(df['eps'], errors='coerce')
        df['np'] = pd.to_numeric(df['np'], errors='coerce')
        df['max_price'] = pd.to_numeric(df['max_price'], errors='coerce')
        df['min_price'] = pd.to_numeric(df['min_price'], errors='coerce')
        df['rd_int'] = df['report_date'].astype(int)
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = int(r['issue_date'])
            pit = df[df['rd_int'] <= iss]
            if len(pit) == 0:
                continue
            win = pit[pit['rd_int'] >= iss - 90]
            if len(win) == 0:
                win = pit.tail(10)
            f = {}
            eps_vals = win['eps'].dropna()
            rating_vals = win['rating_num'].dropna()
            if len(eps_vals) >= 1:
                f['rc_eps_consensus'] = float(eps_vals.mean())
                if len(eps_vals) >= 2:
                    f['rc_eps_dispersion'] = float(eps_vals.std() / (abs(eps_vals.mean()) + 1e-9)) \
                        if abs(eps_vals.mean()) > 1e-9 else None
            if len(eps_vals) >= 2:
                old_eps = eps_vals.iloc[0]
                if abs(old_eps) > 1e-9:
                    f['rc_eps_revision'] = float(eps_vals.iloc[-1] / abs(old_eps) - 1)
            f['rc_analyst_count'] = int(len(win))
            if len(rating_vals) >= 1:
                f['rc_rating_avg'] = float(rating_vals.mean())
                if len(rating_vals) >= 2:
                    f['rc_rating_chg'] = float(rating_vals.iloc[-1] - rating_vals.iloc[0])
            tp_vals = win['max_price'].dropna()
            close = r.get('issue_date_price')
            if len(tp_vals) >= 1 and close and float(close) > 0:
                f['rc_target_upside'] = float(tp_vals.mean() / float(close) - 1)
            if len(eps_vals) >= 2:
                up = int((eps_vals > eps_vals.iloc[0]).sum())
                dn = int((eps_vals < eps_vals.iloc[0]).sum())
                total = up + dn
                if total > 0:
                    f['rc_revision_breadth'] = (up - dn) / total
            f['rc_recency'] = int(iss - int(win['rd_int'].iloc[-1]))
            if f:
                rows.append((stock, str(iss), f))
        if (i + 1) % 500 == 0:
            print(f'  {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)

    print(f'  匹配 {len(rows)} 样本')
    if rows:
        n = batch_update(conn, 'report_rc', rows)
        print(f'  ✅ 回写 {n} 行')
    return len(rows)


def show_status(conn):
    """显示拉取进度。"""
    _ensure_tables(conn)
    fetched = _get_fetched_dates(conn)
    all_dates = _get_trade_dates(conn)
    pending = [d for d in all_dates if d not in fetched]
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM analyst_report_daily")
    total_rows = cur.fetchone()[0]
    cur.execute("SELECT COUNT(*) FROM analyst_report_progress WHERE rows_fetched > 0")
    non_empty = cur.fetchone()[0]
    cur.close()
    print(f'拉取进度: {len(fetched)}/{len(all_dates)} 日 ({len(fetched)/max(len(all_dates),1)*100:.1f}%)')
    print(f'  已拉取: {len(fetched)} 日, 其中 {non_empty} 日有数据')
    print(f'  待拉取: {len(pending)} 日')
    print(f'  总研报: {total_rows} 条')
    if pending:
        print(f'  下一个: {pending[0]}')
    else:
        print('  ✅ 已全部拉取完毕!')


def _sv(v):
    if v is None:
        return None
    try:
        f = float(v)
        return None if (f != f or f == float('inf') or f == float('-inf')) else f
    except (TypeError, ValueError):
        return None


def main():
    ap = argparse.ArgumentParser(description='report_rc 后台慢任务(断点续接)')
    ap.add_argument('action', choices=['fetch', 'compute', 'status'])
    ap.add_argument('--count', type=int, default=1, help='fetch: 拉取几个 report_date(受限于限频,通常1个)')
    ap.add_argument('--limit', type=int, default=0, help='compute: 只算前 N 只股(0=全量)')
    args = ap.parse_args()
    conn = pymysql.connect(**_CFG)
    try:
        if args.action == 'fetch':
            fetch_next(conn, count=args.count)
        elif args.action == 'compute':
            compute_features(conn, limit=args.limit)
        elif args.action == 'status':
            show_status(conn)
    finally:
        conn.close()


if __name__ == '__main__':
    main()
