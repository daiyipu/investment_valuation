#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""统一原始数据摄入(L0): 股票清单 → 逐股 ingest_stock_full(market+FCF+行业) → 共享 DB 表。

**这是任意业务类型(定增/回测/未来)的 L0 唯一入口。** 取数与业务类型无关,
差异只在"喂哪份股票清单"。复用 update_market_data.ingest_stock_full(单股 market+FCF+行业,
参数化 industry_days/skip_*), 不重写单股逻辑。

替代旧 ingest_universe_raw.py(回测专用, 只读 backtest_universe.xlsx); 本脚本接受 resolve_universe
规格, 任意清单(placement/fullA/sample/file)或 Excel 均可, 是其泛化超集。

落同一套共享表(定增/回测都读): market_data / historical_fcf / industry_data / industry_daily。

用法:
  python ingest_raw.py --universe fullA                      # 全A(回测全量)
  python ingest_raw.py --universe sample:500                 # pilot 抽样
  python ingest_raw.py --universe placement                  # 定增股(补 market+FCF+行业; 填定增 batch FCF/行业缺口)
  python ingest_raw.py --src data/backtest_universe.xlsx     # 直接喂 Excel(股票代码 列)
  python ingest_raw.py --universe fullA --skip-market        # 回测: 价量走运行时 pro_bar, 只补 FCF+行业
  python ingest_raw.py --universe fullA --industry-only      # 只重摄行业(修早期缺口, 配合 --industry-days)
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
from update_market_data import ingest_stock_full  # noqa: E402  (复用单股 L0, 不改 update_market_data)
from utils.db_manager import ValuationDB  # noqa: E402
from data_pipeline.fetch_universe import resolve_universe  # noqa: E402


def _resolve_codes(universe, src):
    """股票清单来源: Excel(--src)优先, 否则 resolve_universe(--universe 规格)。"""
    if src:
        df = pd.read_excel(src)
        col = '股票代码' if '股票代码' in df.columns else ('ts_code' if 'ts_code' in df.columns else None)
        if col is None:
            raise ValueError(f'{src} 缺 股票代码/ts_code 列')
        return [str(c).strip() for c in df[col].dropna().unique()]
    return resolve_universe(universe)['ts_code'].astype(str).tolist()


def main():
    ap = argparse.ArgumentParser(description='统一 L0 原始数据摄入(任意业务类型, 复用 ingest_stock_full)')
    ap.add_argument('--universe', default=None,
                    help='placement/fullA/sample:N/file:path(resolve_universe 规格); 与 --src 二选一')
    ap.add_argument('--src', default=None, help='直接喂 Excel(股票代码 列); 优先于 --universe')
    ap.add_argument('--limit', type=int, default=0, help='只处理前 N 只(0=全部, 冒烟用)')
    ap.add_argument('--skip-market', action='store_true', help='跳过 market_data(回测价量走运行时 pro_bar)')
    ap.add_argument('--skip-fcf', action='store_true', help='跳过 historical_fcf')
    ap.add_argument('--skip-industry', action='store_true', help='跳过 industry_data/daily')
    ap.add_argument('--industry-only', action='store_true',
                    help='= --skip-market --skip-fcf(只补行业; 配合 --industry-days 修早期缺口)')
    ap.add_argument('--industry-days', type=int, default=4000,
                    help='行业日线回溯天数(默认4000→start≈2004, 覆盖2010+回测250d回看)')
    args = ap.parse_args()

    if args.industry_only:
        args.skip_market = args.skip_fcf = True
    if not args.universe and not args.src:
        ap.error('需指定 --universe 或 --src')

    codes = _resolve_codes(args.universe, args.src)
    if args.limit:
        codes = codes[:args.limit]
    what = [k for k, skip in [('market_data', args.skip_market), ('FCF', args.skip_fcf), ('行业', args.skip_industry)]
            if not skip]
    src_desc = args.src or f'universe={args.universe}'
    print(f'[{src_desc}] 摄入 {len(codes)} 股 → {"+".join(what) or "无"} (industry_days={args.industry_days})')

    db = ValuationDB()
    ok = fail = 0
    for i, c in enumerate(codes):
        try:
            ingest_stock_full(c, db=db, skip_market=args.skip_market, skip_fcf=args.skip_fcf,
                              skip_industry=args.skip_industry, industry_days=args.industry_days)
            ok += 1
        except Exception as e:
            fail += 1
            print(f'  ⚠️ {c} 失败: {e}')
        if (i + 1) % 25 == 0:
            print(f'  进度 {i+1}/{len(codes)} (ok={ok} fail={fail})')
        time.sleep(0.3)
    print(f'完成: {ok} 成功 / {fail} 失败 / 共 {len(codes)}')


if __name__ == '__main__':
    main()
