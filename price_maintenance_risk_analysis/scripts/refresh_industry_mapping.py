#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""补 industry_data 的【股票→行业映射】(只补分类, 不动 industry_daily 系列)。

industry_data 是 load_pb_vs_industry 的瓶颈: 早年定增股很多从未被摄入行业分类 →
PB_vs_同行中位 算不出(industry_data 缺 stock_code→index_code)。本脚本对缺映射的股
逐个 pro.index_member_all 取申万 l1/l2/l3, 只插分类列(metrics 留空, PB 不需要)。

**不动 industry_daily 系列**(区别于 ingest_raw --industry-only: 后者会逐股重摄系列、
撞 sw_daily 限频走 AKShare 无 pb → 覆盖掉 refresh_industry_daily 修好的数据)。

用法:
  python refresh_industry_mapping.py --src ml_training/data/features.parquet  # 补训练集缺映射股(重训前置)
  python refresh_industry_mapping.py --universe fullA      # 全A
  python refresh_industry_mapping.py --universe placement --limit 20   # 冒烟
"""
import argparse
import os
import sys
import time
from datetime import datetime

import pandas as pd
import pymysql

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))           # price_maintenance_risk_analysis/
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'scripts'))
sys.path.insert(0, os.path.join(PKG, 'scripts', 'data_pipeline'))

from tushare_token import resolve_tushare_token  # noqa: E402
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
import tushare as ts  # noqa: E402
from samples.fetch_universe import resolve_universe  # noqa: E402

_DB = dict(host='127.0.0.1', port=3306, user='root', password='',
           database='investment_valuation', charset='utf8mb4')

# 只插分类列(metrics 留空; PB_vs_同行中位 只读 stock_code+index_code)
_INS_SQL = """INSERT INTO industry_data
(stock_code, index_code, industry_name, sw_l1_code, sw_l1_name, sw_l2_code, sw_l2_name,
 sw_l3_code, sw_l3_name, analysis_date, data_source)
VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""


def main():
    ap = argparse.ArgumentParser(description='补 industry_data 股票→行业映射(只分类, 不动系列)')
    ap.add_argument('--universe', default=None, help='placement/fullA/sample:N/file:path(与 --src 二选一)')
    ap.add_argument('--src', default=None, help='直接喂 parquet/csv/excel(股票代码 列); 优先于 --universe')
    ap.add_argument('--limit', type=int, default=0, help='只处理前 N 只缺映射股(0=全部, 冒烟用)')
    args = ap.parse_args()
    if not args.universe and not args.src:
        ap.error('需指定 --universe 或 --src')

    if args.src:
        df = pd.read_parquet(args.src) if args.src.endswith('.parquet') else pd.read_excel(args.src)
        col = '股票代码' if '股票代码' in df.columns else ('ts_code' if 'ts_code' in df.columns else 'stock_code')
        codes = [str(c).strip() for c in df[col].dropna().unique()]
    else:
        codes = resolve_universe(args.universe)['ts_code'].astype(str).tolist()
    conn = pymysql.connect(**_DB)
    mapped = set(pd.read_sql('SELECT DISTINCT stock_code FROM industry_data', conn)['stock_code'].astype(str))
    conn.close()
    todo = [c for c in codes if c not in mapped]
    if args.limit:
        todo = todo[:args.limit]
    print(f'universe={args.universe}({len(codes)}) | 已映射 {len(codes)-len(todo)} | 待补映射 {len(todo)}')
    if not todo:
        print('✅ 无待补映射股'); return

    pro = ts.pro_api()
    conn = pymysql.connect(**_DB)
    cur = conn.cursor()
    today = datetime.now().strftime('%Y%m%d')
    ok = nocls = fail = 0
    for i, c in enumerate(todo):
        try:
            df = pro.index_member_all(ts_code=c)
            if df is None or df.empty:
                nocls += 1
                continue
            # 取当前生效的(is_new=Y 优先, 否则最新 in_date)
            if 'is_new' in df.columns:
                newest = df[df['is_new'] == 'Y']
                r = newest.iloc[0] if not newest.empty else df.sort_values('in_date').iloc[-1]
            else:
                r = df.sort_values('in_date').iloc[-1]
            l3 = str(r.get('l3_code') or '').strip()
            if not l3:
                nocls += 1
                continue
            cur.execute(_INS_SQL, (
                c, l3,
                r.get('l3_name'),
                r.get('l1_code'), r.get('l1_name'),
                r.get('l2_code'), r.get('l2_name'),
                r.get('l3_code'), r.get('l3_name'),
                today, 'index_member_all'))
            ok += 1
        except Exception as e:
            fail += 1
            if fail <= 5:
                print(f'  ⚠️ {c} 失败: {e}')
        if (i + 1) % 50 == 0:
            conn.commit()
            print(f'  进度 {i+1}/{len(todo)} (补 {ok} / 无分类 {nocls} / 失败 {fail})')
        time.sleep(0.15)
    conn.commit()
    conn.close()
    print(f'完成: 补映射 {ok} / 无申万分类 {nocls} / 失败 {fail} / 共 {len(todo)}')


if __name__ == '__main__':
    main()
