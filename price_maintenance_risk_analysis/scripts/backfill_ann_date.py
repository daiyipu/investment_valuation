#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
回填 financial_indicators.ann_date / end_date（PIT 回溯所需）。
对每只股票取 tushare fina_indicator，按 report_year=year(end_date) 匹配，UPDATE ann_date/end_date。
fina_indicator.ann_date ≈ 年报公告日，用于 load_financial_ratios 做 point-in-time(ann_date <= 报价日)。

用法: python backfill_ann_date.py [--token T]
"""
import argparse
import time
import pymysql
import tushare as ts


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--token', default='f2380d8761bcbf165f87b85f04ed105b1bdcf8721574562294671265')
    ap.add_argument('--sleep', type=float, default=0.18)
    args = ap.parse_args()

    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    cur = conn.cursor()
    cur.execute('SELECT DISTINCT stock_code FROM financial_indicators WHERE ann_date IS NULL')
    codes = [r[0] for r in cur.fetchall()]
    print(f'待回填 ann_date 的股票: {len(codes)}')

    pro = ts.pro_api(args.token)
    ok = rows = fail = 0
    for i, code in enumerate(codes):
        try:
            df = None
            for attempt in range(3):  # 限频退避
                try:
                    df = pro.fina_indicator(ts_code=code, start_date='20140101', end_date='20251231')
                except Exception:
                    df = None
                if df is not None and len(df) > 0:
                    break
                time.sleep(1.2 * (attempt + 1))
            time.sleep(args.sleep)
            if df is None or len(df) == 0 or 'ann_date' not in df.columns:
                fail += 1
                continue
            # 取年报(end_date 以 1231 结尾), 建 report_year -> (ann_date, end_date)
            ann_map = {}
            for _, r in df.iterrows():
                ed = str(r.get('end_date', ''))
                if not ed.endswith('1231'):
                    continue
                yr = int(ed[:4])
                ad = str(r.get('ann_date', '') or '')
                if ad and ad != 'nan':
                    ann_map[yr] = (ad[:8], ed[:8])
            if not ann_map:
                fail += 1
                continue
            # UPDATE 该股票对应 report_year 的 ann_date/end_date
            for yr, (ad, ed) in ann_map.items():
                cur.execute(
                    'UPDATE financial_indicators SET ann_date=%s, end_date=%s WHERE stock_code=%s AND report_year=%s',
                    (ad, ed, code, yr))
                rows += cur.rowcount
            conn.commit()
            ok += 1
            if (i + 1) % 50 == 0:
                print(f'  进度 {i+1}/{len(codes)} ...')
        except Exception as e:
            fail += 1
            print(f'  {code} 失败: {e}')
    conn.close()
    print(f'完成: {ok}/{len(codes)} 只成功, 更新 {rows} 行, 失败 {fail}')


if __name__ == '__main__':
    main()
