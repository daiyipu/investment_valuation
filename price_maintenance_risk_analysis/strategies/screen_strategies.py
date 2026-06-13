"""屏幕 runner: 全市场 or 标的池, 输出波二/抵抗 CSV 列表。

用法:
  python strategies/screen_strategies.py --date 20240601
  python strategies/screen_strategies.py --date 20240601 --universe 000021.SZ,000032.SZ --limit 50

修正计划占位: resist 的 stock_r 取**个股**日收益(stock_qfq→pct_change), 经
aligned_returns 与行业/大盘对齐(非 ind 占位)。全市场扫描大盘仅取一次。
"""
import argparse
import os
import sys

import numpy as np
import pandas as pd
import pymysql

_HERE = os.path.dirname(os.path.abspath(__file__))           # .../strategies
_PKG_ROOT = os.path.dirname(_HERE)                            # .../price_maintenance_risk_analysis
if _PKG_ROOT not in sys.path:
    sys.path.insert(0, _PKG_ROOT)

from strategies.data_loader import stock_qfq_closes, stock_qfq_df, industry_df, market_df, _align_three
from strategies.wave2 import wave2_signal
from strategies.resist import resist_score


def _all_stocks():
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    try:
        return pd.read_sql('SELECT stock_code FROM stocks', conn)['stock_code'].tolist()
    finally:
        conn.close()


def screen(date, universe=None, out_dir='output', limit=None, maxlen=300):
    """universe=None → 全市场(stocks 表); 否则传股票列表。limit 截断(调试用)。

    输出 {out_dir}/wave2_list_{date}.csv 与 resist_list_{date}.csv。
    """
    os.makedirs(out_dir, exist_ok=True)
    codes = universe or _all_stocks()
    if limit:
        codes = codes[:limit]
    print(f'屏幕 {date}: {len(codes)} 只标的...')

    mdf = market_df(date, maxlen)   # 大盘仅取一次, 复用
    w2_rows, rs_rows = [], []
    for i, code in enumerate(codes):
        try:
            # 波二: 个股 qfq 收盘
            c = stock_qfq_closes(code, date, maxlen)
            if len(c) > 250:
                r = wave2_signal(c)
                if r['trigger']:
                    w2_rows.append({'股票代码': code, 'gain': r['gain'], 'retr': r['retr'],
                                    'breakout': r['breakout'], 'score': r['score']})
            # 抵抗: 个股/行业/大盘 对齐收益(个股收益, 非 ind 占位)
            sdf = stock_qfq_df(code, date, maxlen)
            idf = industry_df(code, date, maxlen)
            sr, kr, mr = _align_three(sdf, idf, mdf, maxlen)
            if len(sr) >= 65:
                r = resist_score(sr, kr, mr)
                if r['trigger']:
                    rs_rows.append({'股票代码': code, 'corr_div_stock': r['corr_div_stock'],
                                    'corr_div_sector': r['corr_div_sector'], 'rel_stock': r['rel_stock'],
                                    'rel_sector': r['rel_sector'], 'score': r['score']})
        except Exception as e:
            print(f'  {code} skip: {e}')
        if (i + 1) % 50 == 0:
            print(f'  进度 {i+1}/{len(codes)} | 波二 {len(w2_rows)} / 抵抗 {len(rs_rows)}')

    w2_df = pd.DataFrame(w2_rows).sort_values('score', ascending=False) if w2_rows else pd.DataFrame()
    rs_df = pd.DataFrame(rs_rows).sort_values('score', ascending=False) if rs_rows else pd.DataFrame()
    w2_df.to_csv(f'{out_dir}/wave2_list_{date}.csv', index=False, encoding='utf-8-sig')
    rs_df.to_csv(f'{out_dir}/resist_list_{date}.csv', index=False, encoding='utf-8-sig')
    print(f'完成: 波二 {len(w2_rows)} / 抵抗 {len(rs_rows)} 只, 已输出 CSV 到 {out_dir}/')
    return w2_df, rs_df


if __name__ == '__main__':
    import warnings
    warnings.filterwarnings('ignore')
    ap = argparse.ArgumentParser(description='波二/抵抗 选股屏幕')
    ap.add_argument('--date', required=True, help='屏幕日 YYYYMMDD')
    ap.add_argument('--universe', default=None, help='逗号分隔股票代码, 缺省全市场')
    ap.add_argument('--limit', type=int, default=None, help='截断标的数(调试)')
    ap.add_argument('--out-dir', default='output')
    args = ap.parse_args()
    uni = [s.strip() for s in args.universe.split(',') if s.strip()] if args.universe else None
    screen(args.date, uni, args.out_dir, args.limit)
