#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
回填 financial_indicators：用 EFAES calculate_financial_ratios 从三张报表算 27 个比率
入库——与定增训练集【同源同算法】(非 fina_indicator 快捷值)。

针对 Excel 里 5 个报表字段(inv_turn/ebit_to_interest/cash_to_liqdebt/
cash_to_liqdebt_withinterest/rd_exp_ratio)为 NULL 的股票做完整重算。

用法:
    python backfill_financial_indicators.py <scored_excel> [--token T] [--years 5]
    python backfill_financial_indicators.py backtest_universe.xlsx --force --start-year 2009 --end-year 2025 --years 16  # 回测全A全历史
"""
import argparse
import os
import sys
import time
import numpy as np
import pandas as pd
import pymysql
import tushare as ts

# 复用 EFAES 的财务比率算法(与定增训练集同源同公式)
EFAES_PATH = '/Users/davy/github/EFAES'
if EFAES_PATH not in sys.path:
    sys.path.insert(0, EFAES_PATH)
from src.core.calculate_financial_ratios import calculate_financial_ratios  # noqa: E402

# tushare token: 走 resolve_tushare_token(不再硬编码; 旧 literal f2380... 已泄漏, 需在 tushare.pro 轮换)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from tushare_token import resolve_tushare_token  # noqa: E402
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())

TABLE_FIELDS = [
    'current_ratio', 'quick_ratio', 'inv_turn', 'ar_turn', 'ca_turn', 'assets_turn',
    'roa', 'npta', 'roe', 'roe_dt', 'netprofit_margin', 'grossprofit_margin',
    'debt_to_assets', 'int_to_talcap', 'debt_to_eqt', 'ebit_to_interest',
    'cash_to_liqdebt', 'cash_to_liqdebt_withinterest', 'rd_exp_ratio',
    'op_yoy', 'ebt_yoy', 'netprofit_yoy', 'dt_netprofit_yoy', 'roe_yoy',
    'tr_yoy', 'or_yoy', 'equity_yoy',
]
# fina_indicator 不提供、必须从报表算的 5 个
NEED_STATEMENT = ['inv_turn', 'ebit_to_interest', 'cash_to_liqdebt',
                  'cash_to_liqdebt_withinterest', 'rd_exp_ratio']


def _fetch_statements(pro, code, start, end):
    """取三张表的年报行，附 report_year 列。"""
    out = {}
    for name, fn in [('balancesheet', pro.balancesheet),
                     ('income', pro.income),
                     ('cashflow', pro.cashflow)]:
        df = pd.DataFrame()
        for attempt in range(3):  # tushare 限频会返回空 → 退避重试
            try:
                df = fn(ts_code=code, start_date=start, end_date=end)
            except Exception:
                df = pd.DataFrame()
            if df is not None and len(df) > 0:
                break
            time.sleep(1.5 * (attempt + 1))
        time.sleep(0.25)
        if df is None or len(df) == 0 or 'end_date' not in df.columns:
            out[name] = pd.DataFrame(columns=['report_year'])
            continue
        ed = df['end_date'].astype(str)
        mask = ed.str.endswith('1231')
        df = df[mask].copy()
        df['report_year'] = ed[mask].str[:4].astype(int).values
        out[name] = df.reset_index(drop=True)
    return out


def main():
    ap = argparse.ArgumentParser(description='回填 financial_indicators(三表算法, 与定增同源)')
    ap.add_argument('excel', help='含 股票代码 列的 Excel')
    ap.add_argument('--token', default=None, help='tushare token(默认走 resolve_tushare_token)')
    ap.add_argument('--years', type=int, default=5)
    ap.add_argument('--start-year', type=int, default=2016)  # 多取几年供平均值(当年+上年)
    ap.add_argument('--end-year', type=int, default=2025)
    ap.add_argument('--force', action='store_true',
                    help='全量重算 Excel 每只(新股建行 + 既有行 DELETE+INSERT 覆盖); '
                         '不加则只补"既有行里 5 报表字段缺失"的股(回测全A新股原本无行, null-filter 会漏掉, 必加)')
    args = ap.parse_args()

    ex = pd.read_excel(args.excel, sheet_name='Sheet1')
    codes = [str(c) for c in ex['股票代码'].dropna().unique()]
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    if args.force:
        # 全量: Excel 每只都重算(下方 INSERT 前先 DELETE, 幂等覆盖)。回测全A必走此路——
        # 新股原本无 financial_indicators 行, 默认 null-filter 只查"既有行缺字段"会整批漏掉。
        missing = codes
        print(f'--force 全量重算: Excel {len(codes)} 只(新股建行 + 既有行覆盖)')
    else:
        # 默认: 只补"既有行里 5 报表字段缺失"的(即只被 fina_indicator 快捷覆盖、缺 inv_turn/rd_exp_ratio)
        cs = ','.join([f"'{c}'" for c in codes])
        null_sql = (f"SELECT DISTINCT stock_code FROM financial_indicators "
                    f"WHERE stock_code IN ({cs}) AND (inv_turn IS NULL OR rd_exp_ratio IS NULL)")
        null_codes = set(pd.read_sql(null_sql, conn)['stock_code'])
        missing = [c for c in codes if c in null_codes]
        print(f'Excel {len(codes)} 只; 缺 5 报表字段待重算: {len(missing)}(回测全A新股请加 --force)')

    pro = ts.pro_api(args.token or os.environ['TUSHARE_TOKEN'])
    cur = conn.cursor()
    sd, ed = f'{args.start_year}0101', f'{args.end_year}1231'
    target_years = list(range(args.end_year - args.years + 1, args.end_year + 1))
    cols = ['stock_code', 'report_year'] + TABLE_FIELDS + ['used_average_values', 'valid_indicator_count', 'ann_date', 'end_date']
    ins_sql = f"INSERT INTO financial_indicators ({', '.join(cols)}) VALUES ({', '.join(['%s']*len(cols))})"

    ok = rows = fail = 0
    for i, code in enumerate(missing):
        try:
            stmts = _fetch_statements(pro, code, sd, ed)
            bs, isc, cf = stmts['balancesheet'], stmts['income'], stmts['cashflow']
            if len(bs) == 0 or len(isc) == 0:
                fail += 1
                continue
            computed = []  # (year, [27 vals], valid_count)
            for yr in target_years:
                try:
                    ratios = calculate_financial_ratios(bs, isc, cf, yr)
                except Exception:
                    continue
                if not ratios:
                    continue
                vals, vc = [], 0
                for f in TABLE_FIELDS:
                    v = ratios.get(f)
                    try:
                        v = float(v)
                    except (TypeError, ValueError):
                        v = None
                    if v is None or v != v or np.isinf(v):
                        v = None
                    else:
                        vc += 1
                    vals.append(v)
                if vc > 0:
                    is_row = isc[isc['report_year'] == yr]
                    ad = ed_ = None
                    if len(is_row) > 0:
                        ad = str(is_row.iloc[0].get('ann_date', ''))[:8]
                        ed_ = str(is_row.iloc[0].get('end_date', ''))[:8]
                        ad = ad if ad.isdigit() else None
                        ed_ = ed_ if ed_.isdigit() else None
                    computed.append((yr, vals, vc, ad, ed_))
            if not computed:
                fail += 1
                continue
            cur.execute('DELETE FROM financial_indicators WHERE stock_code=%s', (code,))
            for yr, vals, vc, ad, ed_ in computed:
                cur.execute(ins_sql, [code, yr] + vals + [0, vc, ad, ed_])
                rows += 1
            ok += 1
            if (i + 1) % 25 == 0:
                conn.commit()
                print(f'  进度 {i+1}/{len(missing)} ...')
        except Exception as e:
            fail += 1
            print(f'  {code} 失败: {e}')
    conn.commit()
    conn.close()
    print(f'完成: {ok}/{len(missing)} 只成功, 插入 {rows} 行, 失败 {fail}')


if __name__ == '__main__':
    main()
