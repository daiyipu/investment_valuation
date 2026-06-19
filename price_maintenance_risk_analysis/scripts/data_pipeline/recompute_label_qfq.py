#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用前复权(qfq)价格重算 placement_evaluation 的多期限收益标签。

每只股票取 qfq 全序列一次，对每个样本按 报价日 与 报价日+N月(各期限) 取最近交易日
qfq 收盘，重算各期限涨跌幅(= (close_Nm - close_issue)/close_issue * 100)，单条 UPDATE
写齐所有期限列，回写 DB。

qfq 已含分红再投 = 真实总回报，且无除权跳点，与 qfq 特征同口径。

用法:
  python recompute_label_qfq.py [--horizons 1,3,6,7,12] [--dry-run]
"""
import argparse
import os
import sys
import time
from datetime import datetime
import pymysql
import pandas as pd
import tushare as ts

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from tushare_token import resolve_tushare_token
from utils.db_manager import ValuationDB
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())

_CFG = ValuationDB.MYSQL_CONFIG


def _add_months(ymd, months):
    """YYYYMMDD 字符串 +N 月 → YYYYMMDD 字符串(粗略, 日溢出截到月末)。"""
    y, m, d = int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8])
    tot = y * 12 + (m - 1) + months
    ny, nm = tot // 12, tot % 12 + 1
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


def _ensure_horizon_columns(conn, horizons):
    """幂等: 缺失的 return_{h}m/price_{h}m 列补上(自愈, 兼容旧 MySQL 无 IF NOT EXISTS)。"""
    need = [f'return_{h}m' for h in horizons] + [f'price_{h}m' for h in horizons]
    with conn.cursor() as cur:
        cur.execute(
            "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
            "WHERE TABLE_SCHEMA=%s AND TABLE_NAME='placement_evaluation'", (_CFG['database'],))
        have = {r[0] for r in cur.fetchall()}
    missing = [c for c in need if c not in have]
    if missing:
        ddl = 'ALTER TABLE placement_evaluation ' + ', '.join(f'ADD COLUMN {c} DOUBLE' for c in missing)
        with conn.cursor() as cur:
            cur.execute(ddl)
        conn.commit()
        print(f'  已补列: {missing}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--horizons', default='1,3,6,7,12',
                    help='期限(月), 逗号分隔, 默认 1,3,6,7,12')
    ap.add_argument('--dry-run', action='store_true', help='只打印不回写')
    args = ap.parse_args()
    horizons = sorted({int(x) for x in args.horizons.split(',') if x.strip()})
    max_h = max(horizons)

    conn = pymysql.connect(**_CFG)
    _ensure_horizon_columns(conn, horizons)
    samples = pd.read_sql(
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND issue_date<>'' AND LENGTH(issue_date)=8", conn)
    samples['issue_date'] = samples['issue_date'].astype(str)
    print(f'样本 {len(samples)} 条 / {samples["stock_code"].nunique()} 只股票 | 期限 {horizons}')

    # UPDATE 语句模板: 写齐所有期限列
    set_clause = ', '.join([f'return_{h}m=%s, price_{h}m=%s' for h in horizons])
    upd_sql = (f'UPDATE placement_evaluation SET {set_clause} '
               'WHERE stock_code=%s AND issue_date=%s')
    n_slots = 2 * len(horizons)

    pro = ts.pro_api(os.environ['TUSHARE_TOKEN'])
    cur = conn.cursor()
    ok = upd = fail = 0
    for stock, grp in samples.groupby('stock_code'):
        stock = str(stock)
        # 取数范围覆盖 issue 与最大期限
        dates = []
        for _, r in grp.iterrows():
            dates.append(r['issue_date'])
            dates.append(_add_months(r['issue_date'], max_h))
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
            cmap = dict(zip(df['trade_date'].astype(str),
                            pd.to_numeric(df['close'], errors='coerce')))
        except Exception as e:
            fail += 1; print(f'  {stock} 取数失败: {e}'); continue

        for _, r in grp.iterrows():
            issue = r['issue_date']
            c_i = _nearest_close(cmap, issue)
            if c_i is None or c_i == 0:
                continue
            params = []
            any_h = False
            for h in horizons:
                c_t = _nearest_close(cmap, _add_months(issue, h))
                if c_t is None:
                    ret = None
                else:
                    ret = round((c_t - c_i) / c_i * 100, 4)
                    any_h = True
                params += [ret, round(c_t, 4) if c_t is not None else None]
            params += [stock, issue]
            if not any_h:
                continue
            ok += 1
            if not args.dry_run:
                assert len(params) == n_slots + 2
                cur.execute(upd_sql, params)
                upd += 1
        if not args.dry_run:
            conn.commit()
        done = ok + fail
        if done and done % 50 == 0:
            print(f'  进度: ok={ok}, updated={upd}, failed_stocks={fail}', flush=True)
    if not args.dry_run:
        conn.commit()
    conn.close()
    print(f'完成: {ok} 样本重算, {upd} 行回写, {fail} 股取数失败 | 期限 {horizons}')


if __name__ == '__main__':
    main()
