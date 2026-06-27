#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按【唯一行业指数】刷新 industry_daily 全历史(去重, 非逐股)+ index_factor_pro(指数技术面)。

industry_daily 按 index_code(行业)存, 全市场仅 ~312 个唯一行业。逐股摄入(ingest_raw
--industry-only)会重复摄同一行业数千次且撞 sw_daily 限频; 本脚本按唯一 index_code 去重,
每个行业只取 1 次 fetch_industry_index_data(days=4000→start≈2004), 覆盖 2010+ 回测全历史。

index_factor_pro(同批指数 + 大盘 000300.SH): 87 个指数技术因子(MACD/KDJ/RSI/BOLL/MA/…),
落时序表 index_factor_pro, 供特征层 PIT 切片(≤回溯日期)做指数技术回溯特征。统一在此脚本摄入。

断点续跑: 跳过已有早年(min trade_date ≤ --cutoff)数据的行业; 中断后重跑只补缺口。

典型场景:
  - 定增 PB 重训前置: industry_daily 多数行业只从 ~2023-09 起(days=500 旧摄)→ 早年 issue_date
    无行业 PB → 本脚本补全历史。
  - 回测全历史: 同理。

用法:
  python refresh_industry_daily.py                    # 刷新所有缺早年数据的行业
  python refresh_industry_daily.py --days 4000        # 回溯天数(默认4000→start≈2004)
  python refresh_industry_daily.py --cutoff 20150101  # min≤此值视为已全(跳过; 默认20150101)
  python refresh_industry_daily.py --limit 20         # 冒烟
"""
import argparse
import os
import sys
import time

import pandas as pd

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))           # price_maintenance_risk_analysis/
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'scripts'))
sys.path.insert(0, os.path.join(PKG, 'scripts', 'data_pipeline'))

from tushare_token import resolve_tushare_token  # noqa: E402
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
from update_market_data import fetch_industry_index_data  # noqa: E402
from utils.db_manager import ValuationDB  # noqa: E402

import pymysql  # noqa: E402


def _f_mv(x):
    """nan/inf→None(pymysql 拒 nan)。"""
    try:
        x = float(x); return None if x != x or x in (float('inf'), float('-inf')) else x
    except Exception:
        return None


def ingest_idx_factor_pro(indices):
    """刷新 index_factor_pro(指数技术面 idx_factor_pro, 87因子)→ 时序表 index_factor_pro。
    indices = 指数码列表(申万行业 + 大盘 000300.SH)。每指数全历史1次(8000行/次)。
    PIT: trade_date≤回溯日期 切片做特征。与 industry_daily 同批指数, 统一在此脚本摄入(不另开脚本)。
    幂等建表(ON DUP KEY 更新); _f_mv 处理 NaN(见 save-nan-silent-drop-bug)。"""
    import tushare as ts
    pro = ts.pro_api()
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    cols = None; ok = 0; t0 = time.time()
    with conn.cursor() as cur:
        for n, ic in enumerate(indices, 1):
            df = None
            for a in range(3):
                try:
                    df = pro.idx_factor_pro(ts_code=ic); break
                except Exception:
                    time.sleep(1.5 * (a + 1))
            if df is None or len(df) == 0:
                continue
            df = df.sort_values('trade_date')
            factor_cols = [c for c in df.columns if c not in ('ts_code', 'trade_date')]
            if cols is None:
                cols = factor_cols
                coldef = ','.join(f'`{c}` DOUBLE' for c in cols)
                cur.execute(f'CREATE TABLE IF NOT EXISTS index_factor_pro (index_code VARCHAR(16) NOT NULL, '
                            f'trade_date CHAR(8) NOT NULL, {coldef}, PRIMARY KEY(index_code,trade_date)) '
                            f'ENGINE=InnoDB DEFAULT CHARSET=utf8mb4')
                print(f'  idx_factor_pro 建表: {len(cols)} 因子列')
            placeholders = ','.join(['%s'] * (2 + len(cols)))
            collist = 'index_code,trade_date,' + ','.join(f'`{c}`' for c in cols)
            upd = ','.join(f'`{c}`=VALUES(`{c}`)' for c in cols)
            sql = f'INSERT INTO index_factor_pro ({collist}) VALUES ({placeholders}) ON DUPLICATE KEY UPDATE {upd}'
            rows = [tuple([ic, str(r['trade_date'])] + [_f_mv(r[c]) for c in cols]) for _, r in df.iterrows()]
            B = 1000
            for i in range(0, len(rows), B):
                cur.executemany(sql, rows[i:i + B])
            conn.commit(); ok += 1
            if n % 50 == 0:
                print(f'  idx_factor_pro {n}/{len(indices)} (ok {ok}) {time.time()-t0:.0f}s')
    conn.close()
    print(f'✅ idx_factor_pro: {ok}/{len(indices)} 指数, {len(cols) if cols else 0} 因子 ({(time.time()-t0)/60:.1f}min)')


def main():
    ap = argparse.ArgumentParser(description='按唯一行业指数刷新 industry_daily 全历史(去重+续跑)')
    ap.add_argument('--days', type=int, default=4000, help='回溯天数(默认4000→start≈2004, 覆盖2010+)')
    ap.add_argument('--cutoff', type=int, default=20150101, help='行业 min trade_date≤此值视为已有全历史, 跳过')
    ap.add_argument('--limit', type=int, default=0, help='只处理前 N 个缺早年数据的行业(0=全部, 冒烟用)')
    args = ap.parse_args()

    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    # 唯一行业 + 代表股(取每个 index_code 下的任意一只 stock_code, sw_l3_name, 供 AKShare 降级反查)
    reps = pd.read_sql(
        "SELECT index_code, sw_l3_name, MIN(stock_code) stock_code "
        "FROM industry_data WHERE index_code IS NOT NULL AND index_code<>'' "
        "GROUP BY index_code, sw_l3_name", conn)
    # 已有早年数据的行业(数值比较 trade_date+0, 避免 varchar 词法比较坑): 这些跳过
    done = pd.read_sql(
        f"SELECT DISTINCT index_code FROM industry_daily WHERE trade_date+0 <= {args.cutoff}", conn)
    conn.close()
    done_set = set(done['index_code'].astype(str))

    # 待刷新 = 未在 done_set 里(无早年数据)
    todo = [(str(r['index_code']), r['sw_l3_name'], r['stock_code'])
            for _, r in reps.iterrows() if str(r['index_code']) not in done_set]
    if args.limit:
        todo = todo[:args.limit]

    print(f'唯一行业 {len(reps)} | 已有早年(trade_date≤{args.cutoff})跳过 {len(done_set)} | 待刷新 {len(todo)} (days={args.days})')
    if not todo:
        print('✅ 无待刷新行业'); return

    db = ValuationDB()
    ok = fail = 0
    for i, (ic, name, rep_stock) in enumerate(todo):
        try:
            df = fetch_industry_index_data(ic, days=args.days, stock_code=rep_stock, sw_industry_name=name)
            if df is not None and len(df) >= 250:
                # NaN→None(pymysql 不接受 NaN; save_industry_daily 逐行 row.get 会取到 NaN)
                df = df.replace({float('nan'): None})
                # data_source: tushare_sw 有 pe/pb; AKShare 降级无 pb → 标 akshare_ths
                src = 'tushare_sw'
                pb_col = df.get('pb')
                if pb_col is None or pb_col.dropna().empty:
                    src = 'akshare_ths'
                db.save_industry_daily(ic, df, data_source=src)
                ok += 1
                rng = f"{df['trade_date'].iloc[0]}~{df['trade_date'].iloc[-1]}"
                flag = '' if src == 'tushare_sw' else ' ⚠️AKShare无pb'
                print(f'  [{i+1}/{len(todo)}] {ic} {name}: {len(df)}行 {rng}{flag}')
            else:
                fail += 1
                print(f'  [{i+1}/{len(todo)}] {ic} {name}: 数据不足({0 if df is None else len(df)}行)')
        except Exception as e:
            fail += 1
            print(f'  [{i+1}/{len(todo)}] {ic} {name} 失败: {e}')
        if (i + 1) % 25 == 0:
            print(f'  --- 进度 {i+1}/{len(todo)} (ok={ok} fail={fail})')
        time.sleep(0.3)
    print(f'完成: {ok} 行业刷新成功 / {fail} 失败 / 共 {len(todo)}')

    # 同批指数刷新 index_factor_pro(指数技术面; 大盘+申万行业; 统一在此脚本, 不另开)
    all_idx = sorted({str(r['index_code']) for _, r in reps.iterrows()}) + ['000300.SH']
    print(f'\n刷新 index_factor_pro(指数技术面): {len(all_idx)} 指数(大盘+申万行业)')
    ingest_idx_factor_pro(all_idx)


if __name__ == '__main__':
    main()
