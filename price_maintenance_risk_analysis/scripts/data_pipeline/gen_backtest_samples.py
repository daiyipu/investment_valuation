#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成回测样本清单: (股票代码, 报价日=月末) × universe × 月份, PIT 过滤(在市/非ST/非次新/退市前纳入)。

样本清单是全A多空回测的【单一输入】, 喂给现有管线(零改造):
  - 特征 panel: load_db_features([(code,报价日)...]) + derive_alpha_beta_factors + ... → features_backtest.parquet
  - 7m 前瞻收益: bench_return(报价日, 7m)  (compute_labels)
  - 原始数据摄入: 唯一股去重 → Excel 喂 backfill_financial_indicators / batch_financial_score

即"虚拟报价日"思路: 定增 vs 全A 唯一区别是报价日真假; 取数/算特征只认 (股票, 回溯日期)。
清单里 报价日=月末 即虚拟报价日, 现有管线照跑。

用法:
  python gen_backtest_samples.py --universe sample:500 --years 2010-2025   # pilot
  python gen_backtest_samples.py --universe fullA    --years 2010-2025     # 全A
"""
import argparse
import calendar
import os
import sys

import pandas as pd

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # price_maintenance_risk_analysis/
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'scripts'))
from data_pipeline.fetch_universe import fetch_stock_basic, fetch_namechange, in_universe_at, resolve_universe

DATA_DIR = os.path.join(PKG, 'ml_training', 'data')
SAMPLES_PARQ = os.path.join(DATA_DIR, 'backtest_samples.parquet')
UNIVERSE_XLSX = os.path.join(DATA_DIR, 'backtest_universe.xlsx')


def month_end_dates(year_start, year_end):
    """月末日历日(YYYYMMDD); 特征/收益用 _nearest_close 自动吸附最近交易日(PIT 安全)。"""
    out = []
    for y in range(year_start, year_end + 1):
        for m in range(1, 13):
            out.append(f'{y}{m:02d}{calendar.monthrange(y, m)[1]:02d}')
    return out


def _load_placement_dates():
    """加载定增报价日 → {stock: [issue_date_int, ...] 升序}。
    供回测排除"同时期"样本(训练标签窗口重叠 → IC/L-S 虚高失真)。"""
    import pymysql
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    df = pd.read_sql("SELECT stock_code, issue_date FROM placement_evaluation "
                     "WHERE issue_date IS NOT NULL AND issue_date<>'' AND LENGTH(issue_date)=8", conn)
    conn.close()
    out = {}
    for c, g in df.groupby('stock_code'):
        out[str(c)] = sorted(int(str(d)) for d in g['issue_date'])
    return out


def _is_contemporaneous(stock, date_yyyymmdd, placements, window_months=7):
    """(stock, date) 是否与某定增报价日同时期(|月差|≤window) → 7m 标签窗口重叠 → 排除免泄漏。"""
    dates = placements.get(str(stock))
    if not dates:
        return False
    try:
        y, m = int(str(date_yyyymmdd)[:4]), int(str(date_yyyymmdd)[4:6])
    except (ValueError, IndexError):
        return False
    for idate in dates:
        iy, im = int(str(idate)[:4]), int(str(idate)[4:6])
        if abs((y - iy) * 12 + (m - im)) <= window_months:
            return True
    return False


def main():
    ap = argparse.ArgumentParser(description='生成回测样本清单(股票×月末报价日, PIT过滤)')
    ap.add_argument('--universe', default='sample:500', help='placement/fullA/sample:N/file:path')
    ap.add_argument('--years', default='2010-2025')
    ap.add_argument('--min-list-years', type=int, default=1, help='上市满N年(排次新)')
    ap.add_argument('--exclude-window-months', type=int, default=7,
                    help='排除定增同时期样本的窗口(±N月; 7m标签窗口重叠会致IC/L-S虚高失真)')
    ap.add_argument('--no-exclude-placement', action='store_true',
                    help='不排除定增同时期样本(调试用; 默认排除)')
    args = ap.parse_args()
    ys, ye = map(int, args.years.split('-'))

    # 1) 股票清单 + PIT 元数据(list_date/delist_date/name)
    if args.universe == 'placement':
        stocks = fetch_stock_basic()
    else:
        stocks = resolve_universe(args.universe)
        if 'list_date' not in stocks.columns:           # placement/file 可能缺, 补全A元数据
            full = fetch_stock_basic()
            stocks = stocks.merge(full[['ts_code', 'list_date', 'delist_date', 'name']],
                                  on='ts_code', how='left')
    stocks['ts_code'] = stocks['ts_code'].astype(str)
    stocks['list_date'] = stocks.get('list_date', pd.Series(dtype=str)).fillna('').astype(str)
    stocks['delist_date'] = stocks.get('delist_date', pd.Series(dtype=str)).fillna('').astype(str)
    stocks['name'] = stocks.get('name', stocks['ts_code']).fillna(stocks['ts_code']).astype(str)
    nc = fetch_namechange()

    dates = month_end_dates(ys, ye)
    print(f'universe={args.universe}({len(stocks)}只) × {len(dates)}月末({dates[0]}~{dates[-1]}) → PIT 过滤')

    # 2) PIT 过滤: 每个月末截面, 取在市/非ST/上市≥N年/退市前纳入的股票
    #    并排除定增同时期样本(同股有定增报价日在 ±window 月内 → 7m 标签窗口重叠 → 训练泄漏失真)
    placements = {} if args.no_exclude_placement else _load_placement_dates()
    rows = []
    skipped_leak = 0
    for di, d in enumerate(dates):
        for _, r in stocks.iterrows():
            if not in_universe_at(r, d, nc, min_list_years=args.min_list_years):
                continue
            if placements and _is_contemporaneous(r['ts_code'], d, placements, args.exclude_window_months):
                skipped_leak += 1
                continue
            rows.append({'股票代码': r['ts_code'], '报价日': d})
        if (di + 1) % 24 == 0:
            print(f'  日期 {di+1}/{len(dates)}: 累计 {len(rows)} 样本 (排除泄漏 {skipped_leak})')
    df = pd.DataFrame(rows)
    if skipped_leak:
        print(f'  ⚠️ 排除定增同时期样本 ±{args.exclude_window_months}月: {skipped_leak} 条(避免训练标签窗口重叠失真)')

    # 3) 输出: 全样本 parquet(panel 用) + 唯一股 Excel(喂原始数据摄入脚本)
    df.to_parquet(SAMPLES_PARQ, index=False)
    uniq = df[['股票代码']].drop_duplicates().reset_index(drop=True)
    uniq['股票简称'] = uniq['股票代码']
    uniq['报价日'] = dates[-1]   # 占位: 摄入脚本按 report_year 取数, 报价日仅影响 batch_financial_score 的 score_years 窗口
    uniq.to_excel(UNIVERSE_XLSX, index=False)

    per_date = len(df) / max(1, len(dates))
    print(f'\n✅ 样本清单: {len(df)} 条(股票×月末, 均截面 {per_date:.0f} 只) → {SAMPLES_PARQ}')
    print(f'✅ 唯一股: {len(uniq)} 只 → {UNIVERSE_XLSX}(喂 backfill_financial_indicators / batch_financial_score)')


if __name__ == '__main__':
    main()
