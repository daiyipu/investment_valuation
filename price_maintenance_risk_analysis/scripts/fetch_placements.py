#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抓取近10年A股定向增发名单 + 日期(东方财富 datacenter RPT_SEO_DETAIL)。

数据源: 东方财富数据中心"增发" (reportName=RPT_SEO_DETAIL, 全部增发 5800+条)。
近10年: 增发上市日 >= 2015-01-01; 定向增发(SEO_TYPE=1)。

输出: data/placements_近10年_<date>.xlsx (定向增发名单 + 增发上市日/发行价/募资/发行对象/定价原则等)
用法: python scripts/fetch_placements.py
"""
import os, sys, time, json
import pandas as pd

API = "https://datacenter-web.eastmoney.com/api/data/v1/get"
PARAMS = {"reportName": "RPT_SEO_DETAIL", "columns": "ALL",
          "sortColumns": "", "sortTypes": "", "pageSize": 100}

COLS = {  # eastmoney 字段 → 中文
    'SECUCODE': '股票代码', 'SECURITY_NAME_ABBR': '股票简称', 'SEO_TYPE': '增发类型码',
    'ISSUE_WAY': '发行方式', 'ISSUE_DATE': '发行日(报价日)',
    'ISSUE_LISTING_DATE': '增发上市日', 'EQUITY_RECORD_DATE': '股权登记日',
    'ISSUE_PRICE': '发行价', 'ISSUE_NUM': '发行数量(股)',
    'TOTAL_RAISE_FUNDS': '募资总额', 'NET_RAISE_FUNDS': '净募资',
    'ISSUE_OBJECT': '发行对象', 'PRICE_PRINCIPLE': '定价原则',
    'ISSUE_ON_DATE': '询价开始日', 'ISSUE_OFF_DATE': '询价结束日', 'LOT_DATE': '摇号日',
    'ISSUE_SHARE_BEFORE': '发行前股本', 'ISSUE_SHARE_AFTER': '发行后股本',
    'BVPS_BEFORE': '发行前每股净资产', 'BVPS_AFTER': '发行后每股净资产',
}


def fetch_all():
    import requests
    rows, page = [], 1
    while True:
        r = requests.get(API, params={**PARAMS, "pageNumber": page}, timeout=30).json()
        if not r.get("success") or not r.get("result"):
            print(f"  API 返回: {r.get('message')}"); break
        data = r["result"]["data"]
        rows += data
        total = r["result"].get("count", 0)
        print(f"  page {page}: +{len(data)} (累计 {len(rows)}/{total})")
        if len(data) < 100 or len(rows) >= total:
            break
        page += 1; time.sleep(0.3)
    return rows


def main(years=20):
    print(f"抓取东方财富增发数据 (RPT_SEO_DETAIL) | 定向增发 {'全部' if not years else f'近{years}年'}...")
    raw = fetch_all()
    df = pd.DataFrame(raw)
    out = df[list(COLS.keys())].rename(columns=COLS).copy()
    # 增发类型: SEO_TYPE '1'=定向增发, '2'=公开增发 (eastmoney 返回字符串)
    out['增发类型'] = out['增发类型码'].astype(str).map({'1': '定向增发', '2': '公开增发'}).fillna('其他')
    # 日期解析
    for c in ['发行日(报价日)', '增发上市日', '股权登记日', '询价开始日', '询价结束日', '摇号日']:
        out[c] = pd.to_datetime(out[c], errors='coerce')
    # 过滤: 定向增发 + 时间范围
    dx = out[out['增发类型'] == '定向增发'].copy()
    if years:
        cutoff = pd.Timestamp.now() - pd.DateOffset(years=years)
        dx = dx[dx['发行日(报价日)'].fillna(dx['增发上市日']) >= cutoff]
        span = f'近{years}年'
    else:
        span = '全部'
    dx = dx.sort_values('增发上市日', ascending=False).reset_index(drop=True)
    dx.insert(0, '序号', range(1, len(dx) + 1))

    out_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'ml_training', 'data')
    os.makedirs(out_dir, exist_ok=True)
    tag = pd.Timestamp.now().strftime('%Y%m%d')
    path = os.path.join(out_dir, f'placements_{span}_{tag}.xlsx')
    dx.to_excel(path, index=False)

    print(f"\n✅ 全部增发: {len(out)} 条 | 定向增发({span}): {len(dx)} 条")
    if len(dx):
        print(f"   年份分布:\n{dx['增发上市日'].dt.year.value_counts().sort_index().to_string()}")
    print(f"   写出: {path}")
    # 抽样
    print(f"\n抽样(最新5条):")
    show = dx[['股票代码', '股票简称', '增发上市日', '发行价', '募资总额', '发行方式']].head()
    print(show.to_string(index=False))


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(description='抓取A股定向增发名单(东方财富 RPT_SEO_DETAIL)')
    ap.add_argument('--years', type=int, default=20, help='时间跨度(年, 默认20; 0=全部)')
    args = ap.parse_args()
    main(years=args.years if args.years > 0 else None)
