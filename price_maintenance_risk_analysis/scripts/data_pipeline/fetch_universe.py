#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""全 A universe 构建(多空组合回测用)。

职责:
  1. pro.stock_basic 取全部 A 股(上市 L + 退市 D + 暂停 P),含 list_date/delist_date。
  2. pro.namechange 取 ST/*ST 历史 → 供 PIT 排 ST。
  3. PIT 成员判定: 调仓日 D 在市(list_date≤D<delist_date)、上市≥1 年(排次新)、非 ST。
  4. 可选分层抽样(按申万行业, 固定 seed)→ pilot 用。

缓存到 data/universe.parquet(namechange 缓存 data/namechange.parquet),避免重复取数。
不写 DB,纯只读取数 + 落 parquet。

用法:
  python scripts/data_pipeline/fetch_universe.py                 # 拉/刷新 universe
  python scripts/data_pipeline/fetch_universe.py --sample 500    # 额外产分层抽样 500 只
"""
import argparse
import os
import sys
import random

import pandas as pd

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # price_maintenance_risk_analysis/
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'ml_training'))
DATA_DIR = os.path.join(PKG, 'ml_training', 'data')
UNIVERSE_PATH = os.path.join(DATA_DIR, 'universe.parquet')
NAMECHANGE_PATH = os.path.join(DATA_DIR, 'namechange.parquet')


def _pro():
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    return ts.pro_api()


# ─────────────── 统一股票清单解析(原始数据摄入脚本 --universe 的后端) ───────────────
def resolve_universe(spec='placement', data_dir=None):
    """把一个 universe 规格 解析成 DataFrame[ts_code, name, ...]。

    摄入脚本(update_market_data / backfill_financial_indicators / batch_financial_score)
    共用此函数,取代各自写死读 placement_evaluation 或 Excel。

    spec:
      'placement'  → placement_evaluation 去重 stock_code(默认=定增链, 向后兼容)
      'fullA'      → data/universe.parquet(fetch_universe.py 产的全A ~5000)
      'sample:N'   → data/universe_sample_N.parquet(pilot 分层抽样 N 只)
      'file:path'  → 任意 parquet/csv(须含 ts_code 或 stock_code 列)

    返回 DataFrame, 至少含 ts_code + name 列(placement 模式仅这两列;
    fullA/sample 额外带 list_date/delist_date 供 PIT 过滤)。
    """
    import pymysql
    data_dir = data_dir or DATA_DIR
    if spec == 'placement':
        conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                               database='investment_valuation', charset='utf8mb4')
        df = pd.read_sql(
            "SELECT pe.stock_code ts_code, MAX(s.stock_name) name "
            "FROM placement_evaluation pe LEFT JOIN stocks s ON pe.stock_code=s.stock_code "
            "WHERE pe.issue_date IS NOT NULL AND pe.issue_date<>'' GROUP BY pe.stock_code", conn)
        conn.close()
        if 'name' not in df.columns:
            df['name'] = df['ts_code']
        return df
    # 文件类规格
    if spec.startswith('sample:'):
        path = os.path.join(data_dir, f'universe_sample_{spec.split(":", 1)[1]}.parquet')
    elif spec == 'fullA':
        path = os.path.join(data_dir, 'universe.parquet')
    elif spec.startswith('file:'):
        path = spec.split(':', 1)[1]
    else:
        raise ValueError(f"未知 universe 规格: {spec}(支持 placement/fullA/sample:N/file:path)")
    if not os.path.exists(path):
        raise FileNotFoundError(f'universe 文件不存在: {path}（先跑 fetch_universe.py 生成）')
    df = pd.read_csv(path) if path.endswith('.csv') else pd.read_parquet(path)
    if 'ts_code' not in df.columns and 'stock_code' in df.columns:
        df = df.rename(columns={'stock_code': 'ts_code'})
    if 'name' not in df.columns:
        df['name'] = df['ts_code']
    keep = ['ts_code', 'name'] + [c for c in ['list_date', 'delist_date'] if c in df.columns]
    return df[keep]


def fetch_stock_basic():
    """全 A 股(上市+退市+暂停), 含 list_date/delist_date/name。"""
    pro = _pro()
    frames = []
    for status in ('L', 'D', 'P'):   # 上市/退市/暂停 — 退市股退市前纳入(免生存偏差)
        try:
            df = pro.stock_basic(list_status=status,
                                 fields='ts_code,symbol,name,area,industry,list_date,delist_date')
            frames.append(df)
        except Exception as e:
            print(f'  stock_basic list_status={status} 失败: {e}')
    out = pd.concat(frames, ignore_index=True).drop_duplicates('ts_code')
    out['list_date'] = out['list_date'].astype(str)
    out['delist_date'] = out['delist_date'].fillna('').astype(str)
    print(f'  全 A universe: {len(out)} 只(L+D+P)')
    return out


def fetch_namechange():
    """ST/*ST 历史变更(供 PIT 判定某日是否 ST)。缓存。"""
    if os.path.exists(NAMECHANGE_PATH):
        return pd.read_parquet(NAMECHANGE_PATH)
    pro = _pro()
    # namechange 按交易所逐个拉(ts_code 必填或 exchange)
    frames = []
    for ex in ('SSE', 'SZSE', 'BSE'):
        try:
            df = pro.namechange(exchange=ex, fields='ts_code,name,start_date,end_date,change_reason')
            frames.append(df)
        except Exception as e:
            print(f'  namechange exchange={ex} 失败: {e}')
    out = pd.concat(frames, ignore_index=True).dropna(subset=['ts_code']) if frames else pd.DataFrame()
    out.to_parquet(NAMECHANGE_PATH, index=False)
    print(f'  namechange: {len(out)} 条(缓存 {NAMECHANGE_PATH})')
    return out


def is_st_at(namechange_df, ts_code, date_yyyymmdd):
    """该股在 date 是否处于 ST/*ST 状态。date=YYYYMMDD。"""
    if namechange_df is None or namechange_df.empty:
        return False
    sub = namechange_df[namechange_df['ts_code'] == ts_code]
    for _, r in sub.iterrows():
        s = str(r.get('start_date', ''))
        e = str(r.get('end_date', '') or '')
        name = str(r.get('name', ''))
        if s and s <= date_yyyymmdd and (not e or e >= date_yyyymmdd):
            if 'ST' in name or '*ST' in name:
                return True
    # 也查当前 name(无变更记录但名字带 ST)
    return False


def in_universe_at(stock_row, date_yyyymmdd, namechange_df=None, min_list_years=1):
    """PIT 成员判定: 调仓日 D 在市、上市≥min_list_years、非 ST。"""
    ld = str(stock_row.get('list_date', ''))
    dd = str(stock_row.get('delist_date', '') or '')
    if not ld or ld > date_yyyymmdd:
        return False              # 未上市
    if dd and dd < date_yyyymmdd:
        return False              # 已退市(退市日 < D)
    # 上市满 1 年(排次新)
    try:
        ly = int(ld[:4]); dy = int(date_yyyymmdd[:4])
        if dy - ly < min_list_years:
            return False
    except (ValueError, IndexError):
        return False
    # 排 ST
    cur_name = str(stock_row.get('name', ''))
    if 'ST' in cur_name or '*ST' in cur_name:
        return False
    if namechange_df is not None and is_st_at(namechange_df, stock_row['ts_code'], date_yyyymmdd):
        return False
    return True


def universe_at(stocks_df, date_yyyymmdd, namechange_df=None, min_list_years=1):
    """返回调仓日 D 在市的 ts_code 列表(PIT 过滤后)。"""
    mask = stocks_df.apply(lambda r: in_universe_at(r, date_yyyymmdd, namechange_df, min_list_years), axis=1)
    return stocks_df.loc[mask, 'ts_code'].tolist()


def sample_stratified(stocks_df, n, seed=42):
    """按申万行业(industry 列)分层抽样 n 只(代表性 pilot)。industry 缺失归 '其他'。"""
    rng = random.Random(seed)
    df = stocks_df.copy()
    df['industry'] = df.get('industry', pd.Series(index=df.index)).fillna('其他')
    # 按行业比例抽
    picks = []
    for ind, g in df.groupby('industry'):
        k = max(1, round(len(g) / len(df) * n))
        picks += rng.sample(list(g['ts_code']), min(k, len(g)))
    picks = list(dict.fromkeys(picks))[:n]   # 去重截断
    return picks


def main():
    ap = argparse.ArgumentParser(description='全 A universe 构建(多空回测用)')
    ap.add_argument('--refresh', action='store_true', help='强制重新拉取(否则用缓存)')
    ap.add_argument('--sample', type=int, default=0, help='额外产分层抽样 N 只(0=不抽)')
    args = ap.parse_args()

    if args.refresh or not os.path.exists(UNIVERSE_PATH):
        stocks = fetch_stock_basic()
        stocks.to_parquet(UNIVERSE_PATH, index=False)
        print(f'  缓存 universe → {UNIVERSE_PATH}')
    else:
        stocks = pd.read_parquet(UNIVERSE_PATH)
        print(f'  用缓存 universe: {len(stocks)} 只 ({UNIVERSE_PATH})')
    fetch_namechange()   # 缓存 ST 历史
    if args.sample:
        picks = sample_stratified(stocks, args.sample)
        out = os.path.join(DATA_DIR, f'universe_sample_{args.sample}.parquet')
        stocks[stocks['ts_code'].isin(picks)].to_parquet(out, index=False)
        print(f'  分层抽样 {len(picks)} 只 → {out}')


if __name__ == '__main__':
    main()
