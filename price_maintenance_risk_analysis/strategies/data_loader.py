"""PIT 取数: 按 (stock, date) 取 ≤date 的价格/收益序列, 喂给信号函数。

数据源:
- 个股 qfq 收盘: tushare pro_bar(end_date=date, adj='qfq') —— PIT 正确。
  (market_data.price_series 是 --all 重生成的当前数据, 无日期对齐且非历史 PIT, 仅 fallback。)
- 行业日收益: industry_daily(已有 trade_date+close), 按 industry_data.index_code 取
  (实测 industry_daily.index_code 对应 industry_data.index_code=sw_l3 码 850xxx, 非 sw_l2_code)。
- 大盘日收益: 无 DB 日序列表(market_indices 仅 per-locked_date 快照), 走 tushare index_daily。

三序列须按共同 trade_date 对齐、末项对齐 date → 见 aligned_returns()。
"""
import os
import json

import numpy as np
import pandas as pd
import pymysql

# 与 scripts/update_market_data.py:2004 一致: env 优先, 否则用内置 token
_TUSHARE_TOKEN_DEFAULT = 'f2380d8761bcbf165f87b85f04ed105b1bdcf8721574562294671265'
_MARKET_INDEX = '000300.SH'   # 沪深300


def _conn():
    return pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')


def _ts():
    os.environ.setdefault('TUSHARE_TOKEN', _TUSHARE_TOKEN_DEFAULT)
    import tushare as ts
    return ts


def _norm_date(d):
    """统一为 YYYYMMDD(去横线)。"""
    return str(d).replace('-', '').strip()


# ---------- 个股 ----------
def stock_qfq_df(stock_code, date, maxlen=300):
    """个股 ≤date 的 qfq 日收盘 DataFrame[trade_date, close], 升序。

    PIT: tushare pro_bar(end_date=date) 仅返回 ≤date 数据。
    """
    date = _norm_date(date)
    ts = _ts()
    df = ts.pro_bar(ts_code=stock_code, end_date=date, adj='qfq')
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=['trade_date', 'close'])
    df = df[['trade_date', 'close']].copy()
    df['close'] = df['close'].astype(float)
    df = df.sort_values('trade_date').drop_duplicates('trade_date')
    return df.iloc[-maxlen:].reset_index(drop=True)


def stock_qfq_closes(stock_code, date, maxlen=300):
    """个股 ≤date 的 qfq 收盘(oldest→newest, 末项=对齐 date), numpy array。供 wave2。"""
    return stock_qfq_df(stock_code, date, maxlen)['close'].to_numpy(dtype=float)


# ---------- 行业 ----------
def _industry_index_code(stock_code):
    """该股的行业指数码(industry_daily 键)。实测=industry_data.index_code(sw_l3, 850xxx)。"""
    conn = _conn()
    try:
        cur = conn.cursor()
        cur.execute('SELECT index_code FROM industry_data WHERE stock_code=%s', (stock_code,))
        row = cur.fetchone()
        return row[0] if row and row[0] else None
    finally:
        conn.close()


def industry_df(stock_code, date, maxlen=300):
    """该股所属行业指数 ≤date 日收盘 DataFrame[trade_date, close], 升序。"""
    idx = _industry_index_code(stock_code)
    if not idx:
        return pd.DataFrame(columns=['trade_date', 'close'])
    date = _norm_date(date)
    conn = _conn()
    try:
        df = pd.read_sql(
            'SELECT trade_date, close FROM industry_daily '
            'WHERE index_code=%s AND trade_date<=%s ORDER BY trade_date',
            conn, params=(idx, date))
    finally:
        conn.close()
    if df.empty:
        return df
    df['close'] = df['close'].astype(float)
    return df.iloc[-maxlen:].reset_index(drop=True)


def industry_daily_returns(stock_code, date, maxlen=300):
    """行业指数 ≤date 日收益(Series, 末项对齐 date)。"""
    df = industry_df(stock_code, date, maxlen)
    if df.empty:
        return pd.Series(dtype=float)
    return df['close'].pct_change().dropna()


# ---------- 大盘 ----------
def market_df(date, maxlen=300, index_code=_MARKET_INDEX):
    """大盘指数(默认沪深300) ≤date 日收盘 DataFrame[trade_date, close], 升序。

    无 DB 日序列, 走 tushare index_daily。
    """
    date = _norm_date(date)
    ts = _ts()
    pro = ts.pro_api()
    df = pro.index_daily(ts_code=index_code, end_date=date)
    if df is None or len(df) == 0:
        return pd.DataFrame(columns=['trade_date', 'close'])
    df = df[['trade_date', 'close']].copy()
    df['close'] = df['close'].astype(float)
    df = df.sort_values('trade_date').drop_duplicates('trade_date')
    return df.iloc[-maxlen:].reset_index(drop=True)


def market_daily_returns(date, maxlen=300, index_code=_MARKET_INDEX):
    """大盘 ≤date 日收益(Series, 末项对齐 date)。"""
    df = market_df(date, maxlen, index_code)
    if df.empty:
        return pd.Series(dtype=float)
    return df['close'].pct_change().dropna()


# ---------- 三序列对齐(resist 用) ----------
def _align_three(sdf, idf, mdf, maxlen=300):
    """三个 DataFrame[trade_date,close] → (stock_r, sector_r, market_r) 等长 numpy。

    纯对齐(无取数), 供 aligned_returns 与 screen 复用(避免重复取大盘)。
    """
    if sdf.empty or idf.empty or mdf.empty:
        return np.array([]), np.array([]), np.array([])
    m = (sdf.rename(columns={'close': 's'})
         .merge(idf.rename(columns={'close': 'i'}), on='trade_date', how='inner')
         .merge(mdf.rename(columns={'close': 'm'}), on='trade_date', how='inner')
         .sort_values('trade_date'))
    m = m.iloc[-maxlen:].reset_index(drop=True)
    sr = m['s'].pct_change().dropna().to_numpy(dtype=float)
    kr = m['i'].pct_change().dropna().to_numpy(dtype=float)
    mr = m['m'].pct_change().dropna().to_numpy(dtype=float)
    n = min(len(sr), len(kr), len(mr))
    return sr[-n:], kr[-n:], mr[-n:]


def aligned_returns(stock_code, date, maxlen=300, index_code=_MARKET_INDEX):
    """返回 (stock_r, sector_r, market_r) 三个等长 numpy 日收益数组。

    按共同 trade_date inner-join 对齐, 末项对齐 date, 取末段 maxlen。
    stock_r=个股日收益, sector_r=行业日收益, market_r=大盘日收益。
    """
    sdf = stock_qfq_df(stock_code, date, maxlen)
    idf = industry_df(stock_code, date, maxlen)
    mdf = market_df(date, maxlen, index_code)
    return _align_three(sdf, idf, mdf, maxlen)
