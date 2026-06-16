#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
用前复权(qfq)价格重算 placement_evaluation.return_7m / price_7m。

每只股票取 qfq 全序列，对每个样本按 报价日 和 报价日+7月 取最近交易日的 qfq 收盘，
重算 7 月涨跌幅(= (close_7m - close_issue)/close_issue * 100)，回写 DB。
qfq 已含分红再投 = 真实总回报，且无除权跳点，与 qfq 特征同口径。

用法: python recompute_label_qfq.py [--dry-run]
"""
import argparse
import os
import sys
import time
from datetime import datetime
import pymysql
import numpy as np
import pandas as pd
import tushare as ts

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from tushare_token import resolve_tushare_token
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())


def _add_months(ymd, months):
    """YYYYMMDD 字符串 +N 月 → YYYYMMDD 字符串(粗略, 日溢出截到月末)。"""
    y, m, d = int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8])
    tot = y * 12 + (m - 1) + months
    ny, nm = tot // 12, tot % 12 + 1
    # 月末截断
    import calendar
    last = calendar.monthrange(ny, nm)[1]
    return f'{ny:04d}{nm:02d}{min(d, last):02d}'


def _nearest_close(close_map, target, tol_days=10):
    """close_map: {trade_date_str: close}; 找离 target 日期最近(<=tol_days)的收盘。"""
    try:
        t = datetime.strptime(target, '%Y%m%d')
    except ValueError:
        return None
    best_k, best_dd = None, 1e9
    for k, v in close_map.items():
        if v != v:  # NaN
            continue
        try:
            kd = datetime.strptime(k, '%Y%m%d')
        except ValueError:
            continue
        dd = abs((kd - t).days)
        if dd < best_dd:
            best_dd, best_k = dd, k
    if best_k is None or best_dd > tol_days:
        return None
    return float(close_map[best_k])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--dry-run', action='store_true', help='只打印不回写')
    args = ap.parse_args()

    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    samples = pd.read_sql(
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND issue_date<>'' AND LENGTH(issue_date)=8", conn)
    samples['issue_date'] = samples['issue_date'].astype(str)
    print(f'样本 {len(samples)} 条 / {samples["stock_code"].nunique()} 只股票')

    pro = ts.pro_api(os.environ['TUSHARE_TOKEN'])
    cur = conn.cursor()
    ok = upd = fail = 0
    for stock, grp in samples.groupby('stock_code'):
        stock = str(stock)
        # 该股票所有 issue_date 与 +7月 的范围
        dates = []
        for _, r in grp.iterrows():
            dates.append(r['issue_date'])
            dates.append(_add_months(r['issue_date'], 7))
        sd = (min(dates)[:4] + '0101'); ed = (max(dates)[:4] + '1231')
        try:
            df = None
            for attempt in range(3):
                try:
                    df = ts.pro_bar(ts_code=stock, start_date=sd, end_date=ed, adj='qfq')
                except Exception:
                    df = None
                if df is not None and len(df) > 0:
                    break
                time.sleep(1.2 * (attempt + 1))
            time.sleep(0.3)
            if df is None or len(df) == 0:
                fail += 1; continue
            cmap = dict(zip(df['trade_date'].astype(str), pd.to_numeric(df['close'], errors='coerce')))
        except Exception as e:
            fail += 1; print(f'  {stock} 取数失败: {e}'); continue

        for _, r in grp.iterrows():
            issue = r['issue_date']
            tgt = _add_months(issue, 7)
            c_i = _nearest_close(cmap, issue)
            c_t = _nearest_close(cmap, tgt)
            if c_i is None or c_t is None or c_i == 0:
                continue
            ret = (c_t - c_i) / c_i * 100
            ok += 1
            if not args.dry_run:
                cur.execute(
                    'UPDATE placement_evaluation SET return_7m=%s, price_7m=%s WHERE stock_code=%s AND issue_date=%s',
                    (round(ret, 4), round(c_t, 4), stock, issue))
                upd += 1
        if not args.dry_run:
            conn.commit()
        done = ok + fail
        if done and done % 50 == 0:
            print(f'  进度: ok={ok}, updated={upd}, failed_stocks={fail}', flush=True)
    if not args.dry_run:
        conn.commit()
    conn.close()
    print(f'完成: {ok} 样本重算, {upd} 行回写, {fail} 股取数失败')


if __name__ == '__main__':
    main()
