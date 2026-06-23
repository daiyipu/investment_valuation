#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""5 个特殊特征的独立 PIT loader(多空回测 compute_features 用)。

每个 loader 接受 (codes, date) 或 sample_keys, PIT 过滤(year/ann_date/trade_date ≤ date),
返回 DataFrame(index=codes, cols=特征)。复用既有 DB/API 查询模式, **不 join placement_evaluation**。

PIT 规则(年报披露): pit_max = base_year-1 if month≥5 else base_year-2(4/30 年报披露截止)。
"""
import os
import sys

import numpy as np
import pandas as pd
import pymysql

PKG = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'ml_training'))

_DB = dict(host='127.0.0.1', port=3306, user='root', password='',
           database='investment_valuation', charset='utf8mb4')


def _pit_max(date_yyyymmdd):
    """年报 PIT 年份: 月≥5 用上年(已披露), 否则前年。"""
    s = str(date_yyyymmdd)
    by, mo = int(s[:4]), int(s[4:6])
    return by - 1 if mo >= 5 else by - 2


def _conn():
    return pymysql.connect(**_DB)


# ─────────────── 全历史预取缓存(FCF/总分, 避免逐股逐日查 DB) ───────────────
_FCF_CACHE = {}     # code → DataFrame[year, fcf] 全历史(升序)
_SCORE_CACHE = {}   # code → DataFrame[report_year, total_score] 全历史(升序)


def _chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def prefetch_fcf_scores(codes):
    """一次性批量预取全 universe 的 FCF + 总分全历史(各 N 条 SQL, IN 分块 500),
    缓存到 _FCF_CACHE/_SCORE_CACHE。后续 load_fcf_accel/load_total_score_delta2y 只做 PIT 内存切片。
    把"500股×192月×2查询≈19万次"降到"500股×2≈1000次"(预取) + 内存切片(免费)。"""
    codes = [str(c) for c in codes if c is not None]
    conn = _conn()
    for c in [c for c in codes if c not in _FCF_CACHE]:
        pass   # 触发下面批量
    new_fcf = [c for c in codes if c not in _FCF_CACHE]
    new_sc = [c for c in codes if c not in _SCORE_CACHE]
    if new_fcf:
        for chunk in _chunks(new_fcf, 500):
            ph = ','.join([f"'{c}'" for c in chunk])
            df = pd.read_sql(f"SELECT stock_code, year, fcf FROM historical_fcf "
                             f"WHERE stock_code IN ({ph}) AND fcf IS NOT NULL", conn)
            for c, g in df.groupby('stock_code'):
                _FCF_CACHE[c] = g.sort_values('year').reset_index(drop=True)
    if new_sc:
        for chunk in _chunks(new_sc, 500):
            ph = ','.join([f"'{c}'" for c in chunk])
            df = pd.read_sql(f"SELECT stock_code, report_year, total_score FROM company_annual_scores "
                             f"WHERE stock_code IN ({ph}) AND total_score IS NOT NULL", conn)
            for c, g in df.groupby('stock_code'):
                _SCORE_CACHE[c] = g.sort_values('report_year').reset_index(drop=True)
    conn.close()
    print(f'    预取 FCF({len(_FCF_CACHE)}股) + 总分({len(_SCORE_CACHE)}股) 全历史缓存')


# ─────────────── 1. FCF_加速 (historical_fcf PIT 内存切片) ───────────────
def load_fcf_accel(keys):
    """keys=[(code,date)]。用 _FCF_CACHE 全历史, PIT 切 year≤pit_max 取 T/T1/T2 → YoY 加速度。
    需先 prefetch_fcf_scores 预热(回测 run() 开头调一次)。"""
    if not keys:
        return pd.DataFrame(columns=['FCF_加速'])
    out = {}
    for code, date in keys:
        g = _FCF_CACHE.get(code)
        if g is None or g.empty:
            out[code] = np.nan; continue
        sub = g[g['year'] <= _pit_max(date)]
        rs = sub['fcf'].tail(3).values      # 升序, 末3个 = T/T1/T2(最大年)
        if len(rs) >= 2:
            yoy_t = (rs[-1] - rs[-2]) / abs(rs[-2]) if rs[-2] else np.nan
            yoy_t1 = (rs[-2] - rs[-3]) / abs(rs[-3]) if len(rs) >= 3 and rs[-3] else np.nan
            out[code] = (yoy_t - yoy_t1) if yoy_t1 == yoy_t1 else np.nan
        else:
            out[code] = np.nan
    return pd.DataFrame({'FCF_加速': out})


# ─────────────── 2. 总分_delta_2y (company_annual_scores PIT 内存切片) ───────────────
def load_total_score_delta2y(keys):
    """keys=[(code,date)]。用 _SCORE_CACHE 全历史, PIT 切 report_year≤pit_max 取 T-(T-2)。"""
    if not keys:
        return pd.DataFrame(columns=['总分_delta_2y'])
    out = {}
    for code, date in keys:
        g = _SCORE_CACHE.get(code)
        if g is None or g.empty:
            out[code] = np.nan; continue
        sub = g[g['report_year'] <= _pit_max(date)]
        rs = sub['total_score'].tail(3).values
        out[code] = (rs[-1] - rs[-3]) if len(rs) >= 3 else np.nan
    return pd.DataFrame({'总分_delta_2y': out})


# ─────────────── 3. nb_hold_ratio (pro.hk_hold ≤ date) ───────────────
_HK_CACHE = {}


def load_nb_hold(codes, date_yyyymmdd):
    """pro.hk_hold 全量缓存/股, 截 trade_date≤date 取最新 ratio。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    pro = ts.pro_api()
    d = int(str(date_yyyymmdd))
    out = {}
    for c in codes:
        if c not in _HK_CACHE:
            try:
                df = pro.hk_hold(ts_code=c, fields='trade_date,ratio')
                _HK_CACHE[c] = df
            except Exception:
                _HK_CACHE[c] = pd.DataFrame()
        df = _HK_CACHE[c]
        if df is None or df.empty:
            out[c] = np.nan; continue
        sub = df[pd.to_numeric(df['trade_date'], errors='coerce') <= d]
        out[c] = float(sub['ratio'].iloc[-1]) if not sub.empty else np.nan
    return pd.DataFrame({'nb_hold_ratio': out})


# ─────────────── 4. PB_vs同行中位 (industry_daily + daily_basic PIT) ───────────────
def load_pb_vs_industry(codes, date_yyyymmdd):
    """个股 PB / 行业 PB(industry_daily sw_index_pb ≤ date; 个股 PB 走 daily_basic ≤ date)。
    PIT: 都取 ≤date 最近值。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    pro = ts.pro_api()
    conn = _conn()
    # 行业映射 stock→sw_index
    idmap = pd.read_sql('SELECT stock_code, index_code FROM industry_data', conn)
    conn.close()
    s2i = dict(zip(idmap['stock_code'].astype(str), idmap['index_code'].astype(str)))
    # 行业 PB ≤date(批量)
    ind_codes = [s2i[c] for c in codes if c in s2i]
    ind_pb = {}
    if ind_codes:
        conn = _conn()
        ph = ','.join([f"'{x}'" for x in set(ind_codes)])
        idf = pd.read_sql(f"SELECT index_code, trade_date, pb FROM industry_daily "
                          f"WHERE index_code IN ({ph}) AND trade_date <= '{date_yyyymmdd}'", conn)
        conn.close()
        if not idf.empty:
            idf = idf.sort_values(['index_code', 'trade_date'])
            ind_pb = dict(idf.groupby('index_code')['pb'].last())
    out = {}
    for c in codes:
        try:
            db = pro.daily_basic(ts_code=c, end_date=date_yyyymmdd, fields='trade_date,pb')
            stock_pb = float(db['pb'].iloc[-1]) if db is not None and not db.empty else np.nan
            ic = s2i.get(c); ipb = ind_pb.get(ic, np.nan)
            out[c] = (stock_pb / ipb) if stock_pb == stock_pb and ipb else np.nan
        except Exception:
            out[c] = np.nan
    return pd.DataFrame({'PB_vs_同行中位': out})


# ─────────────── 5. sue_beat (forecast/express/income 披露 PIT) ───────────────
def load_sue_beat(keys):
    """sue_beat: 业绩超预期。披露 PIT(ann_date≤date), 复用 fetch_factors.ingest_sue 披露时间线。
    本期桩: 返回 NaN(score_sc 用训练 median 填, 占 1/15, 影响小, 待精细化接入)。"""
    # TODO: 接 fetch_factors SUE 披露时间线(forecast/express/income ann_date≤date)
    return pd.DataFrame({'sue_beat': {c: np.nan for c, _ in keys}})


def load_specials(codes, date_yyyymmdd):
    """一次性加载 5 个特殊特征, 返回 DataFrame(index=codes)。"""
    keys = [(c, date_yyyymmdd) for c in codes]
    frames = [
        load_fcf_accel(keys),
        load_total_score_delta2y(keys),
        load_nb_hold(codes, date_yyyymmdd),
        load_pb_vs_industry(codes, date_yyyymmdd),
        load_sue_beat(keys),
    ]
    df = pd.concat(frames, axis=1)
    df.index.name = 'code'
    return df
