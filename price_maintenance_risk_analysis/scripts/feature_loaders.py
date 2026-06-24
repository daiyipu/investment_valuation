#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""5 个特殊特征的【规范 PIT loader】(任意业务类型: 定增/回测/未来都用这一套)。

每个 loader 接受 (codes, date) 或 sample_keys, PIT 过滤(year/ann_date/trade_date ≤ date),
返回 DataFrame(index=codes, cols=特征)。复用既有 DB/API 查询模式, **不 join placement_evaluation**。

PIT 规则(年报披露): pit_max = base_year-1 if month≥5 else base_year-2(4/30 年报披露截止)。

────────────── 与定增侧(export_features/derive_features/fetch_factors)等价对照 ──────────────
本模块 = 5 特殊特征的唯一 PIT 真源。定增侧历史另有实现(批量快照回写 placement_evaluation 再读),
逐特征等价性(2026-06 核实):

  特征            | 定增侧实现                              | 与本模块数学等价?
  FCF_加速        | export_features.load_db_features + derive_features.derive_fcf_growth_rates | ✅ 等价(tail(3) YoY 加速度, pit_max 同)
  总分_delta_2y   | load_scored_features_from_db + derive_features.derive_financial_score_deltas | ✅ 等价(T-(T-2), pit_max 同)
  nb_hold_ratio   | fetch_factors.ingest_capitalflow 回写 pe → load_scored 读            | ✅ 等价(hk_hold ≤date, iloc[-1])
  sue_beat        | fetch_factors._sue_for_sample                                     | ✅ 共享同一函数(本模块直接 import)
  PB_vs_同行中位  | derive_features.derive_valuation_relative(peer_companies 中位)     | ❌ 不等价(见下)

  PB_vs_同行中位 口径冲突(统一唯一卡点):
    定增训练口径 = 个股PB / peer_companies 同行中位(非 PIT 快照, 覆盖 53%)
    本模块口径   = daily_basic PB / industry_daily sw_index_pb(PIT ≤date, 覆盖 91%)
    生产 SC 模型 v_sc_20260622 的 15 特征含此列 → 回测喂行业 PB 与训练(peer)分布不同(中位1.17 vs 1.70)
    = PB 特征系统性错配。统一前须先定口径(建议采本模块 PIT 行业口径 → 重训定增模型)。
──────────────────────────────────────────────────────────────────────────────
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


def _pro():
    """共享 tushare pro_api(token 走 resolve_tushare_token, 不硬编码)。"""
    import tushare as ts
    from tushare_token import resolve_tushare_token
    os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
    return ts.pro_api()


# ─────────────── 全历史预取缓存(FCF/总分/SUE, 避免逐股逐日查) ───────────────
_FCF_CACHE = {}     # code → DataFrame[year, fcf] 全历史(升序)
_SCORE_CACHE = {}   # code → DataFrame[report_year, total_score] 全历史(升序)
_SUE_CACHE = {}     # code → 披露时间线 DataFrame(forecast/express/income 合并, ann_date 升序) 或 None
# SUE 时间线落盘缓存: 历史披露是 PIT 固定的(永不改变), 落盘后回测重建 panel 免重取(500股≈15min→秒级)。
# 增量(未来新披露)用 refresh=True 重取覆盖; 否则只补磁盘+内存都没有的股。
_SUE_PARQ = os.path.join(PKG, 'ml_training', 'data', 'sue_timelines.parquet')


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


def _load_sue_disk():
    """从 _SUE_PARQ 载入已预取的 SUE 时间线 → {code: DataFrame 或 None}。
    空时间线存为 _has_data=0 占位行 → 载回 None(避免每轮重取空股的 3 次 API)。
    无文件返回 {}。"""
    if not os.path.exists(_SUE_PARQ):
        return {}
    try:
        df = pd.read_parquet(_SUE_PARQ)
    except Exception:
        return {}
    if df is None or df.empty or 'stock_code' not in df.columns:
        return {}
    out = {}
    for c, g in df.groupby('stock_code'):
        c = str(c)
        if int(g['_has_data'].iloc[0]) == 0:
            out[c] = None
        else:
            out[c] = (g[g['_has_data'] == 1]
                      .drop(columns=['stock_code', '_has_data'], errors='ignore')
                      .reset_index(drop=True))
    return out


def _save_sue_disk():
    """把 _SUE_CACHE 落盘 parquet(增量合并: 保留磁盘已有股, 本次新取的同股覆盖)。
    空时间线写 1 行 _has_data=0 占位; 有效时间线每披露 1 行 _has_data=1。"""
    frames = []
    for c, g in _SUE_CACHE.items():
        c = str(c)
        if g is None or (hasattr(g, 'empty') and g.empty):
            frames.append(pd.DataFrame({'stock_code': [c], '_has_data': [0]}))
            continue
        gg = g.copy()
        gg['stock_code'] = c
        gg['_has_data'] = 1
        frames.append(gg)
    if not frames:
        return
    new_df = pd.concat(frames, ignore_index=True)
    new_codes = set(new_df['stock_code'])
    if os.path.exists(_SUE_PARQ):
        try:
            old = pd.read_parquet(_SUE_PARQ)
            old = old[~old['stock_code'].astype(str).isin(new_codes)]
            new_df = pd.concat([old, new_df], ignore_index=True)
        except Exception:
            pass
    new_df.to_parquet(_SUE_PARQ, index=False)


def prefetch_sue_timelines(codes, refresh=False):
    """预取 forecast/express/income 披露时间线(每股 3 次 tushare API), 缓存供 load_sue_beat PIT 切片。
    复用 fetch_factors._build_disclosure_timeline; 无 DB 依赖(纯 API), 任意股可取。

    历史披露是 PIT 固定的 → 落盘 _SUE_PARQ 后, 回测重建 panel 不再重取(省 ~3 API/股)。
    refresh=True: 忽略磁盘, 全量重取并覆盖(增量新披露用)。"""
    from data_pipeline.fetch_factors import _build_disclosure_timeline
    # 1) 先载入磁盘缓存(除非 refresh)
    if refresh:
        _SUE_CACHE.clear()
    elif not _SUE_CACHE:
        _SUE_CACHE.update(_load_sue_disk())
        if _SUE_CACHE:
            ok0 = sum(1 for v in _SUE_CACHE.values() if v is not None)
            print(f'    载入磁盘 SUE 时间线: {len(_SUE_CACHE)} 股({ok0} 有数据)')
    # 2) 只对磁盘+内存都没有的逐股取 API
    new = [str(c) for c in codes if str(c) not in _SUE_CACHE]
    if new:
        pro = _pro()
        fetched = 0
        for c in new:
            try:
                _SUE_CACHE[c] = _build_disclosure_timeline(pro, c)
                fetched += 1
            except Exception:
                _SUE_CACHE[c] = None
        if fetched:
            _save_sue_disk()
            print(f'    落盘 {fetched} 条新 SUE 时间线 → {os.path.basename(_SUE_PARQ)}')
    ok = sum(1 for v in _SUE_CACHE.values() if v is not None)
    print(f'    预取 SUE 披露时间线: {ok}/{len(_SUE_CACHE)} 股有数据'
          + ('(本次重取)' if refresh else f'(磁盘缓存 {sum(1 for v in _load_sue_disk().values() if v is not None)} 有效)'))


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


# ─────────────── 4. PB_vs同行中位 (industry_daily + daily_basic PIT, 规范口径) ───────────────
_PB_STOCK_CACHE = {}    # code → daily_basic[trade_date, pb] 全量(跨截面复用, 每股只取 1 次 API)
_INDPB_CACHE = None     # (s2i 映射, industry_daily 全量 DataFrame) 一次查, 内存切片


def load_pb_vs_industry(keys):
    """keys=[(code, date)]。个股 daily_basic PB / 行业 industry_daily sw_index_pb, 都 PIT ≤date 取最近。
    返回 {(str(code), str(date)): 个股PB/行业PB}。

    **规范口径**(2026-06 定): 生产 SC 模型 PB_vs_同行中位 统一采此 PIT 行业口径, 与回测同源。
    个股 daily_basic 每股全量缓存(定增/回测跨多截面时每股只 1 次 API); industry_daily 全量一次查。"""
    if not keys:
        return {}
    pro = _pro()
    global _INDPB_CACHE
    if _INDPB_CACHE is None:
        conn = _conn()
        idmap = pd.read_sql('SELECT stock_code, index_code FROM industry_data', conn)
        idf = pd.read_sql("SELECT index_code, trade_date, pb FROM industry_daily", conn)
        conn.close()
        idf['td_int'] = pd.to_numeric(idf['trade_date'], errors='coerce')
        idf = idf.sort_values(['index_code', 'td_int'])
        _INDPB_CACHE = (dict(zip(idmap['stock_code'].astype(str), idmap['index_code'].astype(str))), idf)
    s2i, idf = _INDPB_CACHE

    ind_pb_by_date = {}   # date_int → {index_code: pb ≤date 最新}
    out = {}
    for code, date in keys:
        code, d = str(code), str(date)
        try:
            d_int = int(float(d))
        except ValueError:
            d_int = 0
        # 个股 PB(每股全量缓存, 切 ≤date)
        if code not in _PB_STOCK_CACHE:
            try:
                _PB_STOCK_CACHE[code] = pro.daily_basic(ts_code=code, fields='trade_date,pb')
            except Exception:
                _PB_STOCK_CACHE[code] = pd.DataFrame()
        sdf = _PB_STOCK_CACHE[code]
        stock_pb = np.nan
        if sdf is not None and not sdf.empty:
            sub = sdf[pd.to_numeric(sdf['trade_date'], errors='coerce') <= d_int]
            stock_pb = float(sub['pb'].iloc[-1]) if not sub.empty else np.nan
        # 行业 PB ≤date(按 date 缓存)
        ic = s2i.get(code); ipb = np.nan
        if ic and d_int:
            if d_int not in ind_pb_by_date:
                sub = idf[idf['td_int'] <= d_int]
                ind_pb_by_date[d_int] = dict(sub.groupby('index_code')['pb'].last())
            ipb = ind_pb_by_date[d_int].get(ic, np.nan)
        out[(code, d)] = (stock_pb / ipb) if stock_pb == stock_pb and ipb == ipb and ipb else np.nan
    return out


# ─────────────── 5. sue_beat (forecast/express/income 披露 PIT) ───────────────
def load_sue_beat(keys):
    """sue_beat: 超自家指引=(实际/快报净利 - 预告中点)/|预告中点|, 正=超指引。PIT: ann_date≤date 取最近披露。
    复用 fetch_factors._sue_for_sample(tl, date)。需先 prefetch_sue_timelines; 未预取则惰性逐股建(慢)。"""
    from data_pipeline.fetch_factors import _sue_for_sample, _build_disclosure_timeline
    if not keys:
        return pd.DataFrame(columns=['sue_beat'])
    pro = None
    out = {}
    for code, date in keys:
        c = str(code)
        if c not in _SUE_CACHE:                  # 未预取 → 惰性建(逐股 3 API; 批量请先 prefetch)
            pro = pro or _pro()
            try:
                _SUE_CACHE[c] = _build_disclosure_timeline(pro, c)
            except Exception:
                _SUE_CACHE[c] = None
        tl = _SUE_CACHE[c]
        out[c] = _sue_for_sample(tl, str(date)).get('sue_beat', np.nan) if tl is not None else np.nan
    return pd.DataFrame({'sue_beat': out})


def load_specials(codes, date_yyyymmdd):
    """一次性加载 5 个特殊特征, 返回 DataFrame(index=codes)。"""
    d = str(date_yyyymmdd)
    keys = [(c, d) for c in codes]
    pb_map = load_pb_vs_industry(keys)   # {(code, date): ratio}
    pb_series = pd.Series({c: pb_map.get((c, d), np.nan) for c in codes}, name='PB_vs_同行中位')
    frames = [
        load_fcf_accel(keys),
        load_total_score_delta2y(keys),
        load_nb_hold(codes, date_yyyymmdd),
        pd.DataFrame({'PB_vs_同行中位': pb_series}),
        load_sue_beat(keys),
    ]
    df = pd.concat(frames, axis=1)
    df.index.name = 'code'
    return df
