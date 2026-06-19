#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""统一标签原始值计算 → placement_evaluation DB(单一原始源)。

吸收原 2 脚本: excess(超额收益 excess_mkt/ind_{1,3,7}m) / shortterm(短线 return_1w/2w/4w)。
**只写原始值到 DB**, 不写标签列到 parquet——阈值/极性标签由 export_features 统一发射(与 标签_盈利_* 一致)。
return_*m 由 recompute_label_qfq.py(既有)负责, 本脚本不管。

用法:
  python scripts/compute_labels.py {excess|shortterm|all}
"""
import argparse
import bisect
import calendar
import os
import sys
import time
from datetime import datetime

import numpy as np
import pandas as pd
import pymysql

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(PKG, 'ml_training'))
sys.path.insert(0, PKG)
from utils.db_manager import ValuationDB   # noqa: E402
from tushare_token import resolve_tushare_token   # noqa: E402
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
import tushare as ts   # noqa: E402

_CFG = ValuationDB.MYSQL_CONFIG
EXCESS_COLS = [f'excess_{b}_{h}m' for b in ('mkt', 'ind') for h in (1, 3, 7)]
SHORT_COLS = ['return_1w', 'return_2w', 'return_4w']
WINDOWS = {'1w': 5, '2w': 10, '4w': 20}   # 交易日


def _clean(v):
    try:
        f = float(v)
        return None if f != f else f
    except (TypeError, ValueError):
        return None


def add_months(ymd, months):
    y, m, d = int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8])
    tot = y * 12 + (m - 1) + months
    ny, nm = tot // 12, tot % 12 + 1
    last = calendar.monthrange(ny, nm)[1]
    return f'{ny:04d}{nm:02d}{min(d, last):02d}'


def build_series(close_map):
    arr = []
    for k, v in close_map.items():
        try:
            f = float(v)
        except (TypeError, ValueError):
            continue
        if f != f:
            continue
        try:
            arr.append((datetime.strptime(k, '%Y%m%d').toordinal(), f))
        except ValueError:
            continue
    arr.sort()
    return arr


def _nearest(series, target_ord, tol=15):
    if not series:
        return None
    i = bisect.bisect_left(series, (target_ord, -1e18))
    cands = [series[j] for j in (i - 1, i) if 0 <= j < len(series)]
    if not cands:
        return None
    o, c = min(cands, key=lambda x: abs(x[0] - target_ord))
    return (o, c) if abs(o - target_ord) <= tol else None


def bench_return(series, issue, h):
    try:
        t0 = datetime.strptime(issue, '%Y%m%d').toordinal()
        t1 = datetime.strptime(add_months(issue, h), '%Y%m%d').toordinal()
    except ValueError:
        return None
    c0 = _nearest(series, t0); c1 = _nearest(series, t1)
    if c0 is None or c1 is None or c0[1] == 0:
        return None
    return (c1[1] / c0[1] - 1) * 100


# ── excess ──
def ingest_excess(conn):
    cur = conn.cursor(cursorclass=pymysql.cursors.DictCursor)
    cur.execute("SELECT stock_code, issue_date, return_1m, return_3m, return_7m "
                "FROM placement_evaluation WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8")
    samp = pd.DataFrame(cur.fetchall())
    conn.cursor().close()
    samp['issue_date'] = samp['issue_date'].astype(str)
    print(f'  [excess] 样本 {len(samp)}')
    mkt = ts.pro_api().index_daily(ts_code='000300.SH').sort_values('trade_date')
    mkt_series = build_series(dict(zip(mkt['trade_date'].astype(str), pd.to_numeric(mkt['close'], errors='coerce'))))
    conn_plain = pymysql.connect(**_CFG)
    ind = pd.read_sql('SELECT index_code, trade_date, close FROM industry_daily', conn_plain)
    ind_series = {code: build_series(dict(zip(g['trade_date'].astype(str), pd.to_numeric(g['close'], errors='coerce'))))
                  for code, g in ind.groupby('index_code')}
    smap = pd.read_sql('SELECT stock_code, index_code FROM industry_data', conn_plain)
    conn_plain.close()
    stock2ind = dict(zip(smap['stock_code'], smap['index_code']))
    print(f'  [excess] 大盘 {len(mkt_series)}日 | 行业 {len(ind_series)}指数')

    set_c = ', '.join(f'{c}=%s' for c in EXCESS_COLS)
    upd = f'UPDATE placement_evaluation SET {set_c} WHERE stock_code=%s AND issue_date=%s'
    vals = []
    for _, r in samp.iterrows():
        iss = r['issue_date']; out = {}
        for h in (1, 3, 7):
            sret = r.get(f'return_{h}m')
            if pd.isna(sret):
                continue
            sret = float(sret)
            mr = bench_return(mkt_series, iss, h)
            if mr is not None:
                out[f'excess_mkt_{h}m'] = sret - mr
            iser = ind_series.get(stock2ind.get(r['stock_code']))
            if iser is not None:
                ir = bench_return(iser, iss, h)
                if ir is not None:
                    out[f'excess_ind_{h}m'] = sret - ir
        vals.append(tuple(_clean(out.get(c)) for c in EXCESS_COLS) + (r['stock_code'], iss))
    n = conn.cursor().executemany(upd, vals)
    conn.commit()
    print(f'  ✅ 回写 {n} 行 excess')


# ── shortterm ──
def ingest_shortterm(conn, limit=0):
    cur = conn.cursor(cursorclass=pymysql.cursors.DictCursor)
    cur.execute("SELECT stock_code, issue_date FROM placement_evaluation "
                "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8")
    samp = pd.DataFrame(cur.fetchall())
    conn.cursor().close()
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    set_c = ', '.join(f'{c}=%s' for c in SHORT_COLS)
    upd = f'UPDATE placement_evaluation SET {set_c} WHERE stock_code=%s AND issue_date=%s'
    vals = []; done = 0
    for si, stock in enumerate(stocks):
        grp = samp[samp['stock_code'] == stock]
        dates = sorted(grp['issue_date'].tolist())
        sd = min(dates)[:4] + '0101'; ed = max(dates)[:4] + '1231'
        cmap = {}
        for attempt in range(3):
            try:
                d = ts.pro_bar(ts_code=stock, start_date=sd, end_date=ed, adj='qfq')
                if d is not None and len(d):
                    d = d.sort_values('trade_date')
                    cmap = dict(zip(d['trade_date'].astype(str), pd.to_numeric(d['close'], errors='coerce')))
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
        time.sleep(0.3)
        if not cmap:
            continue
        tds = sorted(cmap.keys())
        for _, r in grp.iterrows():
            iss = r['issue_date']
            i0 = bisect.bisect_left(tds, iss)
            if i0 == len(tds):
                continue
            if tds[i0] > iss and i0 > 0:
                i0 -= 1
            c0 = cmap.get(tds[i0])
            if c0 is None or c0 == 0:
                continue
            out = {}
            for tag, n in WINDOWS.items():
                j = i0 + n
                if j < len(tds):
                    c1 = cmap.get(tds[j])
                    if c1 is not None:
                        out[f'return_{tag}'] = (c1 / c0 - 1) * 100
            vals.append(tuple(_clean(out.get(c)) for c in SHORT_COLS) + (stock, iss))
        if (si + 1) % 200 == 0:
            print(f'  [shortterm] {si+1}/{len(stocks)} | {len(vals)} 样本', flush=True)
    n = conn.cursor().executemany(upd, vals)
    conn.commit()
    print(f'  ✅ 回写 {n} 行 shortterm')


SOURCES = {'excess': ingest_excess, 'shortterm': ingest_shortterm}


def ensure_columns(conn, source):
    cols = EXCESS_COLS if source == 'excess' else SHORT_COLS
    cur = conn.cursor()
    cur.execute("""SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
                   WHERE TABLE_SCHEMA=%s AND TABLE_NAME='placement_evaluation'""", (_CFG['database'],))
    have = {r[0] for r in cur.fetchall()}
    miss = [c for c in cols if c not in have]
    if miss:
        cur.execute('ALTER TABLE placement_evaluation ADD COLUMN ' + ', ADD COLUMN '.join(f'{c} DOUBLE' for c in miss))
        conn.commit()
        print(f'  [{source}] 补列: {miss}')


def main():
    ap = argparse.ArgumentParser(description='统一标签原始值 → placement_evaluation DB')
    ap.add_argument('source', choices=list(SOURCES) + ['all'])
    ap.add_argument('--limit', type=int, default=0)
    args = ap.parse_args()
    conn = pymysql.connect(**_CFG)
    targets = list(SOURCES) if args.source == 'all' else [args.source]
    for src in targets:
        print(f'\n=== {src} ===')
        ensure_columns(conn, src)
        if src == 'shortterm':
            ingest_shortterm(conn, args.limit)
        else:
            SOURCES[src](conn)
    conn.close()


if __name__ == '__main__':
    main()
