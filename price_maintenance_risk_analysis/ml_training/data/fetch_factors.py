#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""统一因子摄入 → placement_evaluation DB(单一原始源) + 时序表。

吸收原 4 个脚本: placement(定增结构) / chip(筹码 cyq_chips) / capitalflow(资金流+北向) / smc(聪明钱 OHLCV)。
全部写 placement_evaluation(key=stock_code+issue_date, 幂等 ADD COLUMN)。SMC 由原"写 parquet"改为写 DB。

★ 时序表迁移(Phase 0.5):
  chip → chip_daily(每日汇总5指标: winner_rate/avg_cost/concentration/peak_price/cost_spread)
  moneyflow → moneyflow_daily(每日原始大小单买卖量额)
  hk_hold → nb_hold_daily(每日北向持股比例+数量)
  ingest_chip/ingest_capitalflow 同时写时序表 + PE 列(过渡期双写, PE 列保留向后兼容)。

用法:
  python scripts/fetch_factors.py {placement|chip|capitalflow|smc|sue|margin|report_rc|pledge|dividend|repurchase|top_list|block_trade|holdernumber|holdertrade|surv|regime|macro|all} [--write] [--limit N]
  无 --write 只 dry-run 打印覆盖。
"""
import argparse
import glob
import os
import re
import sys
import time
from collections import defaultdict

import numpy as np
import pandas as pd
import pymysql

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(PKG, 'ml_training'))
sys.path.insert(0, PKG)
from utils.db_manager import ValuationDB   # noqa: E402
from tushare_token import resolve_tushare_token   # noqa: E402
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
import tushare as ts   # noqa: E402

_CFG = ValuationDB.MYSQL_CONFIG
MASTER = os.path.expanduser(
    '~/产业基金风控/投资情况登记/定增投资登记/易米基金-泛定增类资产总表（周更）2026-06-05.xlsx')

# ── 各源 DB 列定义 (列名, DDL) ──
COLS = {
    'placement': [('em_issue_price', 'DOUBLE'), ('em_issue_num', 'DOUBLE'),
                  ('em_raise_total', 'DOUBLE'), ('em_share_before', 'DOUBLE'),
                  ('em_share_after', 'DOUBLE'), ('em_issue_object', 'VARCHAR(512)'),
                  ('em_price_principle', 'VARCHAR(32)'), ('em_issue_way', 'VARCHAR(32)'),
                  ('pp_unlock_date', 'VARCHAR(8)'), ('pp_underwriter', 'VARCHAR(128)')],
    'chip': [('chip_winner_rate', 'DOUBLE'), ('chip_avg_cost_dev', 'DOUBLE'),
             ('chip_concentration', 'DOUBLE'), ('chip_peak_dev', 'DOUBLE'),
             ('chip_cost_spread', 'DOUBLE')],
    'capitalflow': [('mf_main_net_ratio_5d', 'DOUBLE'), ('mf_main_net_ratio_20d', 'DOUBLE'),
                    ('mf_net_mf_ratio_20d', 'DOUBLE'), ('mf_main_mom', 'DOUBLE'),
                    ('mf_sm_net_ratio_20d', 'DOUBLE'), ('nb_hold_ratio', 'DOUBLE'),
                    ('nb_hold_chg_20d', 'DOUBLE'), ('nb_hold_chg_60d', 'DOUBLE')],
    'smc': [(c, 'DOUBLE') for c in
            ['smc_premium_discount', 'smc_fvg_net', 'smc_bos', 'smc_liq_sweep',
             'smc_displacement', 'smc_ob_retest', 'smc_ote', 'smc_liqvoid']
            + [f'{k}{t}' for t in ('_W', '_M') for k in
               ['smc_premium_discount', 'smc_fvg_net', 'smc_bos', 'smc_liq_sweep',
                'smc_displacement', 'smc_ob_retest', 'smc_ote', 'smc_liqvoid']]],
    'sue': [('sue_yoy', 'DOUBLE'), ('sue_zscore', 'DOUBLE'),
            ('sue_beat', 'DOUBLE'), ('sue_recency_d', 'DOUBLE'),
            ('sue_yoy_mean3', 'DOUBLE'), ('sue_yoy_acc', 'DOUBLE'),
            ('sue_pos_streak', 'DOUBLE'), ('sue_up_trend', 'DOUBLE')],
    # ── P0 新增 ──
    'margin': [('margin_rzye', 'DOUBLE'), ('margin_rqye', 'DOUBLE'),
               ('margin_rzjme_5d', 'DOUBLE'), ('margin_rzjme_20d', 'DOUBLE'),
               ('margin_rzye_chg_20d', 'DOUBLE'), ('margin_rqjme_20d', 'DOUBLE'),
               ('margin_rzrq_balance', 'DOUBLE'), ('margin_rzye_trend_60d', 'DOUBLE')],
    'report_rc': [('rc_eps_consensus', 'DOUBLE'), ('rc_eps_revision', 'DOUBLE'),
                  ('rc_analyst_count', 'INT'), ('rc_rating_avg', 'DOUBLE'),
                  ('rc_rating_chg', 'DOUBLE'), ('rc_target_upside', 'DOUBLE'),
                  ('rc_eps_dispersion', 'DOUBLE'), ('rc_revision_breadth', 'DOUBLE'),
                  ('rc_imp_profit_yoy', 'DOUBLE'), ('rc_recency', 'INT')],
    # ── P1 新增 ──
    'pledge': [('pledge_ratio', 'DOUBLE'), ('pledge_ratio_chg', 'DOUBLE'),
               ('pledge_count', 'INT'), ('pledge_danger_zone', 'INT')],
    'dividend': [('div_yield_ttm', 'DOUBLE'), ('div_payout_ratio', 'DOUBLE'),
                 ('div_yield_chg', 'DOUBLE'), ('div_bonus_shares', 'DOUBLE'),
                 ('div_consistency', 'INT')],
    'repurchase': [('repurchase_amount_ratio', 'DOUBLE'), ('repurchase_recent_d', 'INT'),
                   ('repurchase_count_1y', 'INT'), ('repurchase_progress', 'DOUBLE')],
    # ── P2 新增 ──
    'top_list': [('toplist_count_30d', 'INT'), ('toplist_inst_net_buy', 'DOUBLE'),
                 ('toplist_institutional', 'INT')],
    'block_trade': [('block_count_30d', 'INT'), ('block_discount_avg', 'DOUBLE'),
                    ('block_amount_ratio', 'DOUBLE')],
    'holdernumber': [('holder_count', 'INT'), ('holder_count_chg', 'DOUBLE')],
    'holdertrade': [('insider_net_buy_90d', 'DOUBLE'), ('insider_buy_count_90d', 'INT'),
                    ('insider_direction', 'DOUBLE')],
    'surv': [('surv_count_90d', 'INT'), ('surv_recency', 'INT')],
    # ── P3 新增 ──
    'macro': [('macro_cpi_yoy', 'DOUBLE'), ('macro_ppi_yoy', 'DOUBLE'),
              ('macro_ppi_cpi_spread', 'DOUBLE'), ('macro_pmi', 'DOUBLE'),
              ('macro_pmi_expansion', 'INT'), ('macro_m1_yoy', 'DOUBLE'),
              ('macro_m2_yoy', 'DOUBLE'), ('macro_m1_m2_scissor', 'DOUBLE'),
              ('macro_sf_yoy', 'DOUBLE'), ('macro_shibor_3m', 'DOUBLE'),
              ('macro_us_10y', 'DOUBLE'), ('macro_us_cn_spread', 'DOUBLE'),
              ('macro_lpr_1y', 'DOUBLE'), ('macro_lpr_chg', 'DOUBLE'),
              ('macro_hsgt_net_5d', 'DOUBLE'), ('macro_hsgt_net_20d', 'DOUBLE')],
}


# ── 通用工具 ──
def tcode(s):
    s = str(s).strip().upper()
    if s.endswith(('.SZ', '.SH', '.BJ')):
        return s
    d = re.sub(r'\D', '', s)[:6]
    if len(d) < 6:
        return None
    if d.startswith(('60', '68', '90')):
        return d + '.SH'
    return d + '.SZ'


def ymd(s):
    if pd.isna(s):
        return None
    try:
        return pd.Timestamp(s).strftime('%Y%m%d')
    except Exception:
        d = re.sub(r'\D', '', str(s))
        return d if len(d) == 8 else None


def _clean(v):
    if v is None:
        return None
    try:
        f = float(v)
        return None if f != f else f   # NaN→NULL
    except (TypeError, ValueError):
        return None


def _sv(v):
    """SQL-safe value: NaN/Inf → None。"""
    if v is None:
        return None
    try:
        f = float(v)
        return None if (f != f or f == float('inf') or f == float('-inf')) else f
    except (TypeError, ValueError):
        return None


# ── 时序表写入工具 ──
def _upsert_chip_daily(conn, stock, chip_map):
    """将 cyq_chips 每日筹码分布汇总 → chip_daily 表。
    chip_map: {trade_date_str: [(price, percent), ...]}
    每日汇总5列: winner_rate/avg_cost/concentration/peak_price/cost_spread(绝对值,非相对)。
    """
    if not chip_map:
        return 0
    rows = []
    for td, hist in chip_map.items():
        if not hist:
            continue
        prices = np.array([p for p, _ in hist], float)
        pcts = np.array([q for _, q in hist], float)
        if prices.size == 0 or pcts.sum() <= 0:
            continue
        pcts = pcts / pcts.sum() * 100.0
        wr = float(pcts[prices < prices.mean()].sum() / 100.0)  # 占位winner_rate(无close)
        avg_cost = float((prices * pcts).sum() / 100.0)
        hhi = float(((pcts / 100.0) ** 2).sum())
        peak_price = float(prices[np.argmax(pcts)])
        order = np.argsort(prices)
        ps, qs = prices[order], pcts[order]
        cum = np.cumsum(qs)
        p25 = float(ps[np.searchsorted(cum, 25)])
        p75 = float(ps[np.searchsorted(cum, 75)])
        cost_spread = p75 - p25
        rows.append((td, stock, wr, avg_cost, hhi, peak_price, cost_spread))
    if not rows:
        return 0
    cur = conn.cursor()
    cur.executemany(
        "INSERT INTO chip_daily (trade_date, ts_code, winner_rate, avg_cost, "
        "concentration, peak_price, cost_spread) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s) "
        "ON DUPLICATE KEY UPDATE winner_rate=VALUES(winner_rate), avg_cost=VALUES(avg_cost), "
        "concentration=VALUES(concentration), peak_price=VALUES(peak_price), cost_spread=VALUES(cost_spread)",
        rows)
    conn.commit()
    return len(rows)


def _upsert_moneyflow_daily(conn, stock, mf_df):
    """moneyflow DataFrame → moneyflow_daily 表(原始大小单买卖量额)。"""
    if mf_df is None or len(mf_df) == 0:
        return 0
    d = mf_df.copy()
    rows = []
    for _, r in d.iterrows():
        td = str(r.get('trade_date', ''))
        if len(td) != 8:
            continue
        rows.append((td, stock,
                     _sv(r.get('buy_sm_vol')), _sv(r.get('buy_sm_amount')),
                     _sv(r.get('sell_sm_vol')), _sv(r.get('sell_sm_amount')),
                     _sv(r.get('buy_md_vol')), _sv(r.get('buy_md_amount')),
                     _sv(r.get('sell_md_vol')), _sv(r.get('sell_md_amount')),
                     _sv(r.get('buy_lg_vol')), _sv(r.get('buy_lg_amount')),
                     _sv(r.get('sell_lg_vol')), _sv(r.get('sell_lg_amount')),
                     _sv(r.get('buy_elg_vol')), _sv(r.get('buy_elg_amount')),
                     _sv(r.get('sell_elg_vol')), _sv(r.get('sell_elg_amount')),
                     _sv(r.get('net_mf_vol')), _sv(r.get('net_mf_amount'))))
    if not rows:
        return 0
    cur = conn.cursor()
    cur.executemany(
        "INSERT INTO moneyflow_daily (trade_date, ts_code, "
        "buy_sm_vol, buy_sm_amt, sell_sm_vol, sell_sm_amt, "
        "buy_md_vol, buy_md_amt, sell_md_vol, sell_md_amt, "
        "buy_lg_vol, buy_lg_amt, sell_lg_vol, sell_lg_amt, "
        "buy_elg_vol, buy_elg_amt, sell_elg_vol, sell_elg_amt, "
        "net_mf_vol, net_mf_amt) "
        "VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
        "ON DUPLICATE KEY UPDATE "
        "buy_sm_vol=VALUES(buy_sm_vol), buy_sm_amt=VALUES(buy_sm_amt), "
        "sell_sm_vol=VALUES(sell_sm_vol), sell_sm_amt=VALUES(sell_sm_amt), "
        "buy_md_vol=VALUES(buy_md_vol), buy_md_amt=VALUES(buy_md_amt), "
        "sell_md_vol=VALUES(sell_md_vol), sell_md_amt=VALUES(sell_md_amt), "
        "buy_lg_vol=VALUES(buy_lg_vol), buy_lg_amt=VALUES(buy_lg_amt), "
        "sell_lg_vol=VALUES(sell_lg_vol), sell_lg_amt=VALUES(sell_lg_amt), "
        "buy_elg_vol=VALUES(buy_elg_vol), buy_elg_amt=VALUES(buy_elg_amt), "
        "sell_elg_vol=VALUES(sell_elg_vol), sell_elg_amt=VALUES(sell_elg_amt), "
        "net_mf_vol=VALUES(net_mf_vol), net_mf_amt=VALUES(net_mf_amt)",
        rows)
    conn.commit()
    return len(rows)


def _upsert_nb_hold_daily(conn, stock, nb_df):
    """hk_hold DataFrame → nb_hold_daily 表(每日北向持股比例+数量)。"""
    if nb_df is None or len(nb_df) == 0:
        return 0
    rows = []
    for _, r in nb_df.iterrows():
        td = str(r.get('trade_date', ''))
        if len(td) != 8:
            continue
        ratio = _sv(r.get('ratio'))
        qty = _sv(r.get('vol', r.get('hold_qty', r.get('in_hold'))))
        if ratio is None and qty is None:
            continue
        rows.append((td, stock, ratio, qty))
    if not rows:
        return 0
    cur = conn.cursor()
    cur.executemany(
        "INSERT INTO nb_hold_daily (trade_date, ts_code, hold_ratio, hold_qty) "
        "VALUES (%s,%s,%s,%s) "
        "ON DUPLICATE KEY UPDATE hold_ratio=VALUES(hold_ratio), hold_qty=VALUES(hold_qty)",
        rows)
    conn.commit()
    return len(rows)


def load_sample_keys(conn, since=None):
    """→ [(stock_code, issue_date_str)]。since: 仅 issue_date>=该年(如 '2013')。"""
    cur = conn.cursor()
    sql = "SELECT stock_code, issue_date FROM placement_evaluation WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"
    if since:
        sql += f" AND issue_date >= '{since}'"
    cur.execute(sql)
    return [(r[0], str(r[1])) for r in cur.fetchall()]


def ensure_columns(conn, source):
    cur = conn.cursor()
    cur.execute("""SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
                   WHERE TABLE_SCHEMA=%s AND TABLE_NAME='placement_evaluation'""", (_CFG['database'],))
    have = {r[0] for r in cur.fetchall()}
    miss = [(c, d) for c, d in COLS[source] if c not in have]
    if miss:
        cur.execute('ALTER TABLE placement_evaluation ADD COLUMN ' +
                    ', ADD COLUMN '.join(f'{c} {d}' for c, d in miss))
        conn.commit()
        print(f'  [{source}] 补列: {[c for c, _ in miss]}')


def batch_update(conn, source, rows):
    """rows: [(stock, issue_date, {col: val}), ...]。NaN→NULL。返回 affected。"""
    cols = [c for c, _ in COLS[source]]
    set_c = ', '.join(f'{c}=%s' for c in cols)
    upd = f'UPDATE placement_evaluation SET {set_c} WHERE stock_code=%s AND issue_date=%s'
    vals = [tuple(_clean(r[2].get(c)) for c in cols) + (r[0], r[1]) for r in rows]
    if not vals:
        return 0
    cur = conn.cursor()
    n = cur.executemany(upd, vals)
    conn.commit()
    return n


# ── placement: 东方财富 RPT_SEO_DETAIL + 易米主表 ──
def build_placement(limit=0):
    fs = sorted(glob.glob(os.path.join(PKG, 'ml_training', 'data', 'placements_*.xlsx')))
    if not fs:
        raise FileNotFoundError('无 placements_*.xlsx, 先跑 fetch_placements.py')
    em = pd.read_excel(fs[-1])
    em['tc'] = em['股票代码'].map(tcode)
    em['qd'] = em['发行日(报价日)'].map(ymd)
    em = em.dropna(subset=['tc', 'qd']).reset_index(drop=True)
    em_map = {}
    em_cn = {'em_issue_price': '发行价', 'em_issue_num': '发行数量(股)', 'em_raise_total': '募资总额',
             'em_share_before': '发行前股本', 'em_share_after': '发行后股本',
             'em_issue_object': '发行对象', 'em_price_principle': '定价原则', 'em_issue_way': '发行方式'}
    for _, r in em.iterrows():
        d = {}
        for dbcol, cn in em_cn.items():
            v = r.get(cn)
            if dbcol in ('em_issue_object', 'em_price_principle', 'em_issue_way'):
                d[dbcol] = str(v).strip()[:512 if dbcol == 'em_issue_object' else 32] if pd.notna(v) else None
            else:
                try:
                    d[dbcol] = float(v) if pd.notna(v) else None
                except (TypeError, ValueError):
                    d[dbcol] = None
        em_map[(r['tc'], r['qd'])] = d
    master = {}
    if os.path.exists(MASTER):
        m = pd.read_excel(MASTER, sheet_name='已完成发行定增项目')
        m.columns = [re.sub(r'\s+', '', str(c)) for c in m.columns]
        m['tc'] = m['股票代码'].map(tcode); m['qd'] = m['报价日期'].map(ymd)
        m = m.dropna(subset=['tc', 'qd'])
        for _, r in m.iterrows():
            master[(r['tc'], r['qd'])] = {
                'pp_unlock_date': ymd(r.get('解禁日（预计）')),
                'pp_underwriter': str(r.get('承销商')).strip()[:128] if pd.notna(r.get('承销商')) else None,
            }
    return em_map, master


def ingest_placement(conn, write, limit):
    em_map, master = build_placement(limit)
    keys = load_sample_keys(conn)
    rows = []
    for code, idate in keys:
        d = dict(em_map.get((code, idate), {}))
        d.update(master.get((code, idate), {}))
        if d:
            rows.append((code, idate, d))
    print(f'  [placement] 匹配 {len(rows)}/{len(keys)} (东方财富 {len(em_map)} + 主表 {len(master)})')
    if write:
        n = batch_update(conn, 'placement', rows)
        print(f'  ✅ 回写 {n} 行')


# ── chip: cyq_chips 筹码 ──
def compute_chip_features(hist, close):
    if not hist or close is None or close <= 0:
        return None
    prices = np.array([p for p, _ in hist], float)
    pcts = np.array([q for _, q in hist], float)
    if prices.size == 0 or pcts.sum() <= 0:
        return None
    pcts = pcts / pcts.sum() * 100.0
    winner = pcts[prices < close].sum()
    avg_cost = (prices * pcts).sum() / 100.0
    hhi = ((pcts / 100.0) ** 2).sum()
    peak_price = prices[np.argmax(pcts)]
    order = np.argsort(prices); ps, qs = prices[order], pcts[order]
    cum = np.cumsum(qs)
    p25 = ps[np.searchsorted(cum, 25)]; p75 = ps[np.searchsorted(cum, 75)]
    return {'chip_winner_rate': winner / 100.0, 'chip_avg_cost_dev': avg_cost / close - 1,
            'chip_concentration': hhi, 'chip_peak_dev': peak_price / close - 1,
            'chip_cost_spread': (p75 - p25) / close}


def fetch_stock_chips(pro, stock, dates):
    if not dates:
        return {}
    sd, ed = min(dates) - 7, max(dates) + 7
    out = defaultdict(list)
    for attempt in range(3):
        try:
            df = pro.cyq_chips(ts_code=stock, start_date=str(sd), end_date=str(ed))
            if df is not None and len(df):
                for _, r in df.iterrows():
                    out[str(r['trade_date'])].append((float(r['price']), float(r['percent'])))
            return dict(out)
        except Exception:
            time.sleep(1.0 * (attempt + 1))
    return {}


def nearest_td(chip_map, issue_date, tol=10):
    best, bd = None, 1e9
    iid = int(issue_date)
    for td in chip_map:
        try:
            d = abs(int(td) - iid)
        except ValueError:
            continue
        if d < bd:
            bd, best = d, td
    return best if bd <= tol else None


def _dict_query(conn, sql):
    """执行 SELECT → 返回 list[dict](兼容 pymysql 1.4.6 不支持 cursor(cursorclass=...))。"""
    cur = conn.cursor()
    cur.execute(sql)
    cols = [d[0] for d in cur.description]
    rows = [dict(zip(cols, r)) for r in cur.fetchall()]
    cur.close()
    return rows


def ingest_chip(conn, write, limit):
    """筹码分布 → chip_daily 时序表 + placement_evaluation(过渡双写)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date, issue_date_price FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8 AND issue_date >= '2018'"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    ts_rows = 0  # 时序表写入计数
    for i, stock in enumerate(stocks):
        sd = samp[samp['stock_code'] == stock]
        dates = sorted(int(x) for x in sd['issue_date'])
        cm = fetch_stock_chips(pro, stock, dates)
        # ★ 写 chip_daily 时序表
        if write and cm:
            ts_rows += _upsert_chip_daily(conn, stock, cm)
        # PE 兼容写入(过渡期保留)
        for _, r in sd.iterrows():
            td = nearest_td(cm, r['issue_date'])
            cp = r.get('issue_date_price')
            f = compute_chip_features(cm.get(td, []), float(cp) if pd.notna(cp) else None) if td else None
            if f:
                rows.append((stock, r['issue_date'], f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [chip] {i+1}/{len(stocks)} | PE {len(rows)} 样本 | 时序 {ts_rows} 行', flush=True)
    print(f'  [chip] PE 匹配 {len(rows)} 样本 | 时序表 {ts_rows} 行')
    if write:
        n = batch_update(conn, 'chip', rows)
        print(f'  ✅ PE 回写 {n} 行 + 时序表 {ts_rows} 行')


# ── capitalflow: moneyflow + hk_hold ──
def _fetch_series(pro, stock, dates, api):
    if not dates:
        return None
    sd, ed = min(dates) - 70, max(dates) + 7
    fn = {'moneyflow': pro.moneyflow, 'hk_hold': pro.hk_hold}[api]
    for attempt in range(3):
        try:
            df = fn(ts_code=stock, start_date=str(sd), end_date=str(ed))
            if df is not None and len(df):
                return df.sort_values('trade_date').reset_index(drop=True)
            return None
        except Exception:
            time.sleep(1.0 * (attempt + 1))
    return None


def mf_features(mf_df, issue_yd):
    if mf_df is None:
        return {}
    d = mf_df.copy()
    for c in ['buy_elg_amount', 'sell_elg_amount', 'buy_lg_amount', 'sell_lg_amount',
              'buy_sm_amount', 'sell_sm_amount', 'buy_md_amount', 'sell_md_amount', 'net_mf_amount']:
        d[c] = pd.to_numeric(d[c], errors='coerce')
    d['main_net'] = d['buy_elg_amount'] + d['buy_lg_amount'] - d['sell_elg_amount'] - d['sell_lg_amount']
    d['sm_net'] = d['buy_sm_amount'] - d['sell_sm_amount']
    d['total'] = (d['buy_elg_amount'] + d['sell_elg_amount'] + d['buy_lg_amount'] + d['sell_lg_amount']
                  + d['buy_md_amount'] + d['sell_md_amount'] + d['buy_sm_amount'] + d['sell_sm_amount'])
    d = d[d['trade_date'].astype(int) <= issue_yd]
    if len(d) == 0:
        return {}
    d['main_ratio'] = d['main_net'] / d['total'].replace(0, np.nan)
    d['net_ratio'] = d['net_mf_amount'] / d['total'].replace(0, np.nan)
    d['sm_ratio'] = d['sm_net'] / d['total'].replace(0, np.nan)
    l5, l20 = d.tail(5), d.tail(20)
    return {'mf_main_net_ratio_5d': l5['main_ratio'].mean(), 'mf_main_net_ratio_20d': l20['main_ratio'].mean(),
            'mf_net_mf_ratio_20d': l20['net_ratio'].mean(),
            'mf_main_mom': l5['main_ratio'].mean() - l20['main_ratio'].mean(),
            'mf_sm_net_ratio_20d': l20['sm_ratio'].mean()}


def nb_features(nb_df, issue_yd):
    if nb_df is None:
        return {}
    d = nb_df.copy()
    d['ratio'] = pd.to_numeric(d.get('ratio'), errors='coerce')
    d = d.dropna(subset=['ratio'])
    d = d[d['trade_date'].astype(int) <= issue_yd]
    if len(d) == 0:
        return {}
    cur = d['ratio'].iloc[-1]
    out = {'nb_hold_ratio': cur}
    if len(d) > 20:
        out['nb_hold_chg_20d'] = cur - d['ratio'].iloc[-21]
    if len(d) > 60:
        out['nb_hold_chg_60d'] = cur - d['ratio'].iloc[-61]
    return out


def ingest_capitalflow(conn, write, limit):
    """资金流+北向 → moneyflow_daily + nb_hold_daily 时序表 + placement_evaluation(过渡双写)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8 AND issue_date >= '2013'"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    mf_ts, nb_ts = 0, 0  # 时序表写入计数
    for i, stock in enumerate(stocks):
        sd = samp[samp['stock_code'] == stock]
        dates = sorted(int(x) for x in sd['issue_date'])
        mf = _fetch_series(pro, stock, dates, 'moneyflow')
        nb = _fetch_series(pro, stock, dates, 'hk_hold')
        # ★ 写时序表
        if write:
            mf_ts += _upsert_moneyflow_daily(conn, stock, mf)
            nb_ts += _upsert_nb_hold_daily(conn, stock, nb)
        # PE 兼容写入(过渡期保留)
        for _, r in sd.iterrows():
            f = {}
            f.update(mf_features(mf, int(r['issue_date'])))
            f.update(nb_features(nb, int(r['issue_date'])))
            if f:
                rows.append((stock, r['issue_date'], f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [capitalflow] {i+1}/{len(stocks)} | PE {len(rows)} 样本 | mf时序 {mf_ts} nb时序 {nb_ts}', flush=True)
    print(f'  [capitalflow] PE 匹配 {len(rows)} 样本 | 时序表 mf={mf_ts} nb={nb_ts}')
    if write:
        n = batch_update(conn, 'capitalflow', rows)
        print(f'  ✅ PE 回写 {n} 行 + 时序表 mf={mf_ts} nb={nb_ts}')


# ── smc: OHLCV → smc_factors(daily+W/M) ──
def ingest_smc(conn, write, limit):
    from features.derive_features import _get_stock_ohlcv, prefetch_ohlcv
    from features.factor_engine import smc_factors, smc_factors_multiperiod
    base = pd.read_parquet(os.path.join(PKG, 'ml_training', 'data', 'features.parquet'))
    base['_kqd'] = base['报价日'].map(lambda v: str(int(float(v))) if pd.notna(v) else None)
    base = base.dropna(subset=['_kqd']).reset_index(drop=True)
    prefetch_ohlcv(base['股票代码'].astype(str).unique(), max_workers=5)
    stocks = sorted(base['股票代码'].astype(str).unique())
    if limit:
        stocks = stocks[:limit]
    rows = []
    for i, code in enumerate(stocks):
        sub = base[base['股票代码'].astype(str) == code]
        try:
            sd, oh = _get_stock_ohlcv(code)
        except Exception:
            sd = None
        if sd is None:
            continue
        for _, r in sub.iterrows():
            d = r['_kqd']
            m = sd <= d
            if not m.any():
                continue
            sd2 = sd[m]; o = oh['open'][m]; h = oh['high'][m]; l = oh['low'][m]; c = oh['close'][m]
            if len(c) < 25:
                continue
            f = smc_factors(o, h, l, c)
            f.update(smc_factors_multiperiod(sd2, o, h, l, c))
            rows.append((code, d, f))
        if (i + 1) % 500 == 0:
            print(f'  [smc] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [smc] 匹配 {len(rows)} 样本')
    if write:
        n = batch_update(conn, 'smc', rows)
        print(f'  ✅ 回写 {n} 行')


# ── sue: 业绩预告/快报/利润表 → 业绩超预期(同比 surprise + 超自家指引) ──
def _fetch_disc_df(pro, stock, api):
    """forecast/express/income 全历史, 3 重试。"""
    for attempt in range(3):
        try:
            if api == 'forecast':
                return pro.forecast(ts_code=stock)
            if api == 'express':
                return pro.express(ts_code=stock)
            return pro.income(ts_code=stock,
                              fields='ts_code,ann_date,end_date,n_income,n_income_attr_p')
        except Exception:
            time.sleep(1.0 * (attempt + 1))
    return None


def _build_disclosure_timeline(pro, stock):
    """合并 forecast/express/income → 按 ann_date 升序的披露时间线。

    每条: {ann_date, end_date, source, n_income, yoy}。income 同时提供
    prior-year 同期做 yoy 随机游走基准(无分析师一致预期时的"预期")。
    """
    inc = _fetch_disc_df(pro, stock, 'income')
    ex = _fetch_disc_df(pro, stock, 'express')
    fc = _fetch_disc_df(pro, stock, 'forecast')

    # income: end_date → (ann_date, 归母净利)。同期多次披露(修正)取最早(首次)。
    inc_map = {}
    if inc is not None and len(inc):
        inc = inc.dropna(subset=['ann_date', 'end_date'])
        for _, r in inc.iterrows():
            ed, ad = str(r['end_date']), str(r['ann_date'])
            if len(ed) != 8 or len(ad) != 8:
                continue
            ni = _clean(r.get('n_income_attr_p'))
            if ni is None:
                ni = _clean(r.get('n_income'))
            if ni is None:
                continue
            if ed not in inc_map or ad < inc_map[ed][0]:
                inc_map[ed] = (ad, ni)

    def prior_year_end(ed):
        try:
            return str(int(ed[:4]) - 1) + ed[4:]
        except Exception:
            return None

    def yoy_from_baseline(ed, ni):
        py = prior_year_end(ed)
        if py and py in inc_map and inc_map[py][1] not in (0, None):
            return ni / abs(inc_map[py][1]) - 1
        return None

    rows = []
    for ed, (ad, ni) in inc_map.items():
        rows.append({'ann_date': ad, 'end_date': ed, 'source': 'income',
                     'n_income': ni, 'yoy': yoy_from_baseline(ed, ni)})
    if ex is not None and len(ex):
        for _, r in ex.iterrows():
            ed, ad = str(r.get('end_date', '')), str(r.get('ann_date', ''))
            if len(ed) != 8 or len(ad) != 8:
                continue
            ni = _clean(r.get('n_income'))
            yoy = None
            for yf in ('yoy_net_profit', 'n_income_yoy', 'net_profit_yoy'):
                v = _clean(r.get(yf))
                if v is not None:
                    yoy = v / 100.0
                    break
            if yoy is None and ni is not None:
                yoy = yoy_from_baseline(ed, ni)
            rows.append({'ann_date': ad, 'end_date': ed, 'source': 'express',
                         'n_income': ni, 'yoy': yoy})
    if fc is not None and len(fc):
        for _, r in fc.iterrows():
            ed, ad = str(r.get('end_date', '')), str(r.get('ann_date', ''))
            if len(ed) != 8 or len(ad) != 8:
                continue
            # forecast net_profit_min/max 单位为万元; income/express 为元 → 统一换算为元
            nmin, nmax = _clean(r.get('net_profit_min')), _clean(r.get('net_profit_max'))
            if nmin is not None and nmax is not None:
                ni = (nmin + nmax) / 2 * 1e4
            elif nmin is not None:
                ni = nmin * 1e4
            elif nmax is not None:
                ni = nmax * 1e4
            else:
                ni = None
            pmin, pmax = _clean(r.get('p_change_min')), _clean(r.get('p_change_max'))
            if pmin is not None and pmax is not None:
                yoy = (pmin + pmax) / 2 / 100.0
            elif ni is not None:
                yoy = yoy_from_baseline(ed, ni)
            else:
                yoy = None
            rows.append({'ann_date': ad, 'end_date': ed, 'source': 'forecast',
                         'n_income': ni, 'yoy': yoy})

    if not rows:
        return None
    tl = pd.DataFrame(rows).sort_values('ann_date').reset_index(drop=True)
    tl['ad_int'] = tl['ann_date'].astype(int)
    return tl


def _sue_for_sample(tl, issue_date):
    """报价日前最近一期披露 → sue_yoy/zscore/beat/recency。PIT: ann_date ≤ 报价日。"""
    if tl is None or len(tl) == 0:
        return {}
    iss = int(issue_date)
    pit = tl[tl['ad_int'] <= iss]
    if len(pit) == 0:
        return {}
    cur = pit.iloc[-1]
    out = {}
    if pd.notna(cur['yoy']):
        out['sue_yoy'] = float(cur['yoy'])
    try:
        out['sue_recency_d'] = (pd.Timestamp(str(iss)) - pd.Timestamp(cur['ann_date'])).days
    except Exception:
        pass
    # z-score: 该股 income 同比序列的自身历史标准化(Latane-Jones SUE)
    inc_yoy = tl[(tl['source'] == 'income') & tl['yoy'].notna()]['yoy']
    if len(inc_yoy) >= 3 and pd.notna(cur['yoy']):
        s = inc_yoy.std()
        if s and s > 1e-9:
            out['sue_zscore'] = float((cur['yoy'] - inc_yoy.mean()) / s)
    # 超自家指引: 当前(快报/实际) vs 同 end_date 最早 forecast(指引)
    if cur['source'] != 'forecast' and pd.notna(cur['n_income']):
        fc_same = pit[(pit['end_date'] == cur['end_date']) & (pit['source'] == 'forecast')]
        if len(fc_same):
            f0 = fc_same.iloc[0]
            if pd.notna(f0['n_income']) and abs(f0['n_income']) > 1e-9:
                out['sue_beat'] = float((cur['n_income'] - f0['n_income']) / abs(f0['n_income']))
    # 盈利动量/趋势(中长期信号: 报价日前已披露的年报 YoY 序列的趋势/持续性, 补 7m 弱的短板)
    inc_ann = pit[(pit['source'] == 'income') & pit['yoy'].notna()
                  & pit['end_date'].astype(str).str.endswith('1231')]
    if len(inc_ann):
        ys = inc_ann['yoy'].tolist()
        out['sue_yoy_mean3'] = float(np.mean(ys[-3:]))       # 近3年报同比均值(盈利水平, 稳定→中长期)
        if len(ys) >= 2:
            out['sue_yoy_acc'] = float(ys[-1] - ys[-2])      # 最新-上年同比(盈利加速度)
        streak = 0
        for y in reversed(ys):                                # 连续盈利年限(YoY>0, 持续盈利能力)
            if y > 0:
                streak += 1
            else:
                break
        out['sue_pos_streak'] = float(streak)
        if len(ys) >= 3:
            last3 = ys[-3:]
            out['sue_up_trend'] = float(np.mean(
                [1 if last3[i + 1] > last3[i] else 0 for i in range(len(last3) - 1)]))   # 改善趋势比例
    return out


def ingest_sue(conn, write, limit):
    cur = conn.cursor()
    cur.execute("SELECT stock_code, issue_date FROM placement_evaluation "
                "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8")
    samp = pd.DataFrame(cur.fetchall(), columns=['stock_code', 'issue_date'])
    cur.close()
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows, skip = [], 0
    for i, stock in enumerate(stocks):
        tl = _build_disclosure_timeline(pro, stock)
        if tl is None:
            skip += 1
            continue
        for _, r in samp[samp['stock_code'] == stock].iterrows():
            f = _sue_for_sample(tl, r['issue_date'])
            if f:
                rows.append((stock, r['issue_date'], f))
        time.sleep(0.25)
        if (i + 1) % 200 == 0:
            print(f'  [sue] {i+1}/{len(stocks)} | {len(rows)} 样本 (跳过 {skip} 无披露)', flush=True)
    print(f'  [sue] 匹配 {len(rows)} 样本 (跳过 {skip} 只股无披露)')
    if write:
        n = batch_update(conn, 'sue', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P0: margin_detail 融资融券 ──
def _upsert_margin_daily(conn, df):
    """margin_detail DataFrame → margin_daily 表(全市场当日)。"""
    if df is None or len(df) == 0:
        return 0
    rows = []
    for _, r in df.iterrows():
        rows.append((str(r['trade_date']), str(r['ts_code']),
                     _sv(r.get('rzye')), _sv(r.get('rqye')),
                     _sv(r.get('rzmre')), _sv(r.get('rqyl')),
                     _sv(r.get('rzche')), _sv(r.get('rqchl')),
                     _sv(r.get('rqmcl')), _sv(r.get('rzrqye'))))
    if not rows:
        return 0
    cur = conn.cursor()
    cur.executemany(
        "INSERT INTO margin_daily (trade_date, ts_code, rzye, rqye, rzmre, rqyl, "
        "rzche, rqchl, rqmcl, rzrqye) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) "
        "ON DUPLICATE KEY UPDATE rzye=VALUES(rzye), rqye=VALUES(rqye), rzmre=VALUES(rzmre), "
        "rqyl=VALUES(rqyl), rzche=VALUES(rzche), rqchl=VALUES(rqchl), "
        "rqmcl=VALUES(rqmcl), rzrqye=VALUES(rzrqye)", rows)
    conn.commit()
    return len(rows)


def _margin_features_from_ts(margin_df, issue_date):
    """从 margin_daily 切片计算 8 个特征。PIT: trade_date <= issue_date。"""
    if margin_df is None or len(margin_df) == 0:
        return {}
    d = margin_df.copy()
    d = d[d['trade_date'].astype(int) <= int(issue_date)].sort_values('trade_date')
    if len(d) == 0:
        return {}
    rzye = pd.to_numeric(d['rzye'], errors='coerce')
    rqye = pd.to_numeric(d['rqye'], errors='coerce')
    rzche = pd.to_numeric(d['rzche'], errors='coerce')
    rzmre = pd.to_numeric(d['rzmre'], errors='coerce')
    rqmcl = pd.to_numeric(d['rqmcl'], errors='coerce')
    rzjme = rzmre - rzche  # 融资净买入
    rqjme = pd.to_numeric(d['rqmcl'], errors='coerce')  # 融券净卖出(近似)
    out = {}
    cur_rzye = _sv(rzye.iloc[-1])
    cur_rqye = _sv(rqye.iloc[-1])
    if cur_rzye is not None:
        out['margin_rzye'] = cur_rzye
    if cur_rqye is not None:
        out['margin_rqye'] = cur_rqye
    # 5日融资净买入
    if len(rzjme) >= 5:
        out['margin_rzjme_5d'] = _sv(rzjme.iloc[-5:].sum())
    # 20日融资净买入
    if len(rzjme) >= 20:
        out['margin_rzjme_20d'] = _sv(rzjme.iloc[-20:].sum())
    # 融资余额20日变化率
    if len(rzye) >= 21 and rzye.iloc[-21] > 0:
        out['margin_rzye_chg_20d'] = _sv(rzye.iloc[-1] / rzye.iloc[-21] - 1)
    # 20日融券净卖出
    if len(rqjme) >= 20:
        out['margin_rqjme_20d'] = _sv(rqjme.iloc[-20:].sum())
    # 多空力量对比
    if cur_rzye and cur_rqye and (cur_rzye + cur_rqye) > 0:
        out['margin_rzrq_balance'] = _sv((cur_rzye - cur_rqye) / (cur_rzye + cur_rqye))
    # 融资余额60日趋势(slope)
    if len(rzye) >= 60:
        y = rzye.iloc[-60:].values.astype(float)
        x = np.arange(len(y), dtype=float)
        mask = ~np.isnan(y)
        if mask.sum() >= 30:
            slope = np.polyfit(x[mask], y[mask], 1)[0]
            out['margin_rzye_trend_60d'] = _sv(slope / (y[mask].mean() + 1e-9))
    return out


def ingest_margin(conn, write, limit):
    """融资融券 → margin_daily 时序表 + placement_evaluation 特征列。
    按 trade_date 批量拉取(一次全市场 ~3800 股),避免逐股调用。
    """
    # 确定日期范围
    cur = conn.cursor()
    cur.execute("SELECT MIN(issue_date), MAX(issue_date) FROM placement_evaluation "
                "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8")
    min_d, max_d = cur.fetchone()
    cur.close()
    if not min_d:
        print('  [margin] 无样本'); return
    # 检查 margin_daily 已有数据(增量)
    cur = conn.cursor()
    cur.execute("SELECT MAX(trade_date) FROM margin_daily")
    last_td = cur.fetchone()[0]
    cur.close()
    start_d = str(int(last_td) + 1) if last_td else '20100101'
    end_d = str(max_d)
    if start_d > end_d:
        print(f'  [margin] 时序表已是最新(至 {last_td})')
    else:
        # 拉取 trade_date 列表(从交易日历)
        pro = ts.pro_api()
        cal = pro.trade_cal(exchange='SSE', start_date=start_d, end_date=end_d, is_open='1')
        if cal is None or len(cal) == 0:
            print('  [margin] 无法获取交易日历'); return
        tds = sorted(cal['cal_date'].tolist())
        if limit:
            tds = tds[:limit]
        ts_rows = 0
        for i, td in enumerate(tds):
            for attempt in range(3):
                try:
                    df = pro.margin_detail(trade_date=td)
                    if df is not None and len(df):
                        if write:
                            ts_rows += _upsert_margin_daily(conn, df)
                    break
                except Exception:
                    time.sleep(1.5 * (attempt + 1))
            if (i + 1) % 100 == 0:
                print(f'  [margin] {i+1}/{len(tds)} 日 | 时序 {ts_rows} 行', flush=True)
            time.sleep(0.15)
        print(f'  [margin] 时序表新增 {ts_rows} 行')

    # 从 margin_daily 计算 PE 特征
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8 AND issue_date >= '2010'"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    rows = []
    for i, stock in enumerate(stocks):
        # 从 margin_daily 读该股全量
        cur = conn.cursor()
        cur.execute("SELECT trade_date, rzye, rqye, rzmre, rqyl, rzche, rqchl, rqmcl, rzrqye "
                    "FROM margin_daily WHERE ts_code=%s ORDER BY trade_date", (stock,))
        cols = [d[0] for d in cur.description]
        mdata = pd.DataFrame(cur.fetchall(), columns=cols)
        cur.close()
        if len(mdata) == 0:
            continue
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            f = _margin_features_from_ts(mdata, r['issue_date'])
            if f:
                rows.append((stock, r['issue_date'], f))
        if (i + 1) % 500 == 0:
            print(f'  [margin] PE 特征 {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [margin] PE 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'margin')
        n = batch_update(conn, 'margin', rows)
        print(f'  ✅ PE 回写 {n} 行')


# ── P0: report_rc 券商盈利预测 ──
_RATING_MAP = {'买入': 5, '强烈推荐': 5, '推荐': 5, '增持': 4, '谨慎推荐': 4,
               '中性': 3, '持有': 3, '减持': 2, '卖出': 1, '回避': 1}


def ingest_report_rc(conn, write, limit):
    """券商盈利预测 → placement_evaluation 特征列(直写 PE)。
    按 ts_code 拉取全历史研报,PIT: report_date <= issue_date。
    """
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date, issue_date_price FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        for attempt in range(3):
            try:
                df = pro.report_rc(ts_code=stock)
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
                df = None
        if df is None or len(df) == 0:
            continue
        # 预处理: 数值化 + 评级映射
        df['eps'] = pd.to_numeric(df['eps'], errors='coerce')
        df['np'] = pd.to_numeric(df['np'], errors='coerce')
        df['max_price'] = pd.to_numeric(df['max_price'], errors='coerce')
        df['min_price'] = pd.to_numeric(df['min_price'], errors='coerce')
        df['rating_num'] = df['rating'].map(_RATING_MAP)
        df['rd_int'] = pd.to_numeric(df['report_date'], errors='coerce').astype('Int64')
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = int(r['issue_date'])
            pit = df[df['rd_int'] <= iss].dropna(subset=['rd_int'])
            if len(pit) == 0:
                continue
            # 近90日窗口
            win = pit[pit['rd_int'] >= iss - 90]
            if len(win) == 0:
                win = pit.tail(10)  # fallback: 最近10条
            f = {}
            eps_vals = win['eps'].dropna()
            np_vals = win['np'].dropna()
            rating_vals = win['rating_num'].dropna()
            # 一致预期 EPS
            if len(eps_vals) >= 1:
                f['rc_eps_consensus'] = float(eps_vals.mean())
                if len(eps_vals) >= 2:
                    f['rc_eps_dispersion'] = _sv(float(eps_vals.std() / (abs(eps_vals.mean()) + 1e-9)))
            # EPS 修正(vs 90日窗口最早)
            if len(eps_vals) >= 2:
                old_eps = eps_vals.iloc[0]
                new_eps = eps_vals.iloc[-1]
                if abs(old_eps) > 1e-9:
                    f['rc_eps_revision'] = _sv(float(new_eps / abs(old_eps) - 1))
            # 分析师覆盖数
            f['rc_analyst_count'] = int(len(win))
            # 一致评级
            if len(rating_vals) >= 1:
                f['rc_rating_avg'] = float(rating_vals.mean())
                if len(rating_vals) >= 2:
                    f['rc_rating_chg'] = _sv(float(rating_vals.iloc[-1] - rating_vals.iloc[0]))
            # 目标价隐含上涨
            tp_vals = win['max_price'].dropna()
            close = r.get('issue_date_price')
            if len(tp_vals) >= 1 and pd.notna(close) and float(close) > 0:
                avg_tp = float(tp_vals.mean())
                f['rc_target_upside'] = _sv(avg_tp / float(close) - 1)
            # 修正广度
            if len(eps_vals) >= 2:
                up = int((eps_vals > eps_vals.iloc[0]).sum())
                dn = int((eps_vals < eps_vals.iloc[0]).sum())
                total = up + dn
                if total > 0:
                    f['rc_revision_breadth'] = _sv((up - dn) / total)
            # 预测利润同比
            if len(np_vals) >= 1:
                f['rc_imp_profit_yoy'] = _sv(float(np_vals.iloc[-1]))  # 绝对值占位
            # 信息新鲜度
            f['rc_recency'] = int(iss - int(win['rd_int'].iloc[-1]))
            if f:
                rows.append((stock, r['issue_date'], f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [report_rc] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [report_rc] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'report_rc')
        n = batch_update(conn, 'report_rc', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P1: pledge_stat 股权质押 ──
def ingest_pledge(conn, write, limit):
    """股权质押 → placement_evaluation(直写 PE)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        for attempt in range(3):
            try:
                df = pro.pledge_stat(ts_code=stock)
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
                df = None
        if df is None or len(df) == 0:
            continue
        df['end_date'] = df['end_date'].astype(str)
        df = df.sort_values('end_date')
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = r['issue_date']
            pit = df[df['end_date'] <= iss]
            if len(pit) == 0:
                continue
            cur = pit.iloc[-1]
            f = {}
            ratio = _sv(cur.get('pledge_ratio'))
            if ratio is not None:
                f['pledge_ratio'] = ratio / 100.0  # 百分比→比例
                f['pledge_danger_zone'] = 1 if ratio > 50 else 0
            count = _sv(cur.get('pledge_count'))
            if count is not None:
                f['pledge_count'] = int(count)
            if len(pit) >= 2:
                prev_ratio = _sv(pit.iloc[-2].get('pledge_ratio'))
                if prev_ratio is not None and ratio is not None:
                    f['pledge_ratio_chg'] = (ratio - prev_ratio) / 100.0
            if f:
                rows.append((stock, iss, f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [pledge] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [pledge] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'pledge')
        n = batch_update(conn, 'pledge', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P1: dividend 分红送股 ──
def ingest_dividend(conn, write, limit):
    """分红送股 → placement_evaluation(直写 PE)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date, issue_date_price FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        for attempt in range(3):
            try:
                df = pro.dividend(ts_code=stock)
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
                df = None
        if df is None or len(df) == 0:
            continue
        df['ann_date'] = df.get('ann_date', pd.Series()).astype(str)
        df['stk_div'] = pd.to_numeric(df.get('stk_div', pd.Series()), errors='coerce').fillna(0)
        df['cash_div'] = pd.to_numeric(df.get('cash_div', pd.Series()), errors='coerce').fillna(0)
        df['stk_bo_rate'] = pd.to_numeric(df.get('stk_bo_rate', pd.Series()), errors='coerce').fillna(0)
        df['stk_co_rate'] = pd.to_numeric(df.get('stk_co_rate', pd.Series()), errors='coerce').fillna(0)
        # 用现金分红+送股合计;只取已实施(div_proc=实施)
        df_impl = df[df.get('div_proc', pd.Series()) == '实施'].copy() if 'div_proc' in df.columns else df.copy()
        if len(df_impl) == 0:
            df_impl = df  # fallback: 不过滤
        df_impl['total_div'] = df_impl['cash_div'] + df_impl['stk_div']
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = r['issue_date']
            pit = df_impl[(df_impl['ann_date'] <= iss) & (df_impl['ann_date'].str.len() == 8)]
            if len(pit) == 0:
                continue
            close = r.get('issue_date_price')
            f = {}
            # TTM 股息率(近4季累计分红/股价)
            if pd.notna(close) and float(close) > 0:
                recent = pit.tail(4)
                total_div = recent['total_div'].sum()
                if total_div > 0:
                    f['div_yield_ttm'] = float(total_div / float(close))
            # 分红率(占盈利比例,此处简化为每股分红/1 — 精确需 EPS)
            latest_div = _sv(pit.iloc[-1].get('stk_div'))
            if latest_div and latest_div > 0:
                f['div_payout_ratio'] = latest_div  # 占位,后续可 /EPS
            # 股息率变化(vs 1年前)
            year_ago = pit[pit['ann_date'] <= str(int(iss[:4]) - 1) + iss[4:]]
            if len(year_ago) > 0 and pd.notna(close) and float(close) > 0:
                old_div = year_ago.tail(4)['total_div'].sum()
                if old_div > 0 and total_div > 0:
                    f['div_yield_chg'] = _sv(float(total_div / old_div - 1))
            # 送转增
            latest_bo = _sv(pit.iloc[-1].get('stk_bo_rate'))
            latest_co = _sv(pit.iloc[-1].get('stk_co_rate'))
            if latest_bo or latest_co:
                f['div_bonus_shares'] = _sv((latest_bo or 0) + (latest_co or 0))
            # 连续分红年数
            years = sorted(set(pit['ann_date'].str[:4].tolist()))
            streak = 0
            for j in range(len(years) - 1, 0, -1):
                if int(years[j]) == int(years[j - 1]) + 1:
                    streak += 1
                else:
                    break
            f['div_consistency'] = streak + 1 if years else 0
            if f:
                rows.append((stock, iss, f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [dividend] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [dividend] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'dividend')
        n = batch_update(conn, 'dividend', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P1: repurchase 股票回购 ──
def ingest_repurchase(conn, write, limit):
    """股票回购 → placement_evaluation(直写 PE)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        for attempt in range(3):
            try:
                df = pro.repurchase(ts_code=stock)
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
                df = None
        if df is None or len(df) == 0:
            continue
        df['ann_date'] = df.get('ann_date', pd.Series()).astype(str)
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = r['issue_date']
            pit = df[(df['ann_date'] <= iss) & (df['ann_date'].str.len() == 8)]
            if len(pit) == 0:
                continue
            f = {}
            # 近1年回购次数
            yr_ago = str(int(iss[:4]) - 1) + iss[4:]
            recent = pit[pit['ann_date'] >= yr_ago]
            f['repurchase_count_1y'] = int(len(recent))
            # 最近回购距报价日天数
            try:
                f['repurchase_recent_d'] = int(
                    (pd.Timestamp(iss) - pd.Timestamp(pit['ann_date'].iloc[-1])).days)
            except Exception:
                pass
            # 回购金额/总市值(简化: 用 repurchase_amount 如有)
            amt = _sv(pit.iloc[-1].get('repurchase_amount', pit.iloc[-1].get('amt')))
            if amt and amt > 0:
                f['repurchase_amount_ratio'] = _sv(amt)  # 占位
            # 完成进度
            proc = _sv(pit.iloc[-1].get('proc', pit.iloc[-1].get('progress')))
            if proc is not None:
                f['repurchase_progress'] = _sv(proc / 100.0 if proc > 1 else proc)
            if f:
                rows.append((stock, iss, f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [repurchase] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [repurchase] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'repurchase')
        n = batch_update(conn, 'repurchase', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P2: top_list + top_inst 龙虎榜 ──
def ingest_top_list(conn, write, limit):
    """龙虎榜 → placement_evaluation(直写 PE)。按 trade_date 批量拉。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8 AND issue_date >= '2010'"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    stock_set = set(stocks)
    # 收集所有样本的报价日前30日范围内的龙虎榜
    all_dates = set()
    for d in samp['issue_date']:
        for offset in range(31):
            try:
                dt = pd.Timestamp(d) - pd.Timedelta(days=offset)
                all_dates.add(dt.strftime('%Y%m%d'))
            except Exception:
                pass
    all_dates = sorted(all_dates)
    pro = ts.pro_api()
    # 按日期拉龙虎榜
    toplist_map = defaultdict(lambda: {'count': 0, 'inst_net': 0.0, 'has_inst': 0})
    for i, td in enumerate(all_dates):
        for attempt in range(3):
            try:
                df = pro.top_list(trade_date=td)
                if df is not None and len(df):
                    for _, r in df.iterrows():
                        ts_c = str(r.get('ts_code', ''))
                        if ts_c in stock_set:
                            toplist_map[(ts_c, td)]['count'] += 1
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
        if (i + 1) % 200 == 0:
            print(f'  [top_list] 日期 {i+1}/{len(all_dates)}', flush=True)
        time.sleep(0.15)
    # 聚合到样本
    rows = []
    for _, r in samp.iterrows():
        code, iss = r['stock_code'], r['issue_date']
        if code not in stock_set:
            continue
        cnt, inst_net, has_inst = 0, 0.0, 0
        for offset in range(31):
            try:
                dt = (pd.Timestamp(iss) - pd.Timedelta(days=offset)).strftime('%Y%m%d')
            except Exception:
                continue
            entry = toplist_map.get((code, dt))
            if entry:
                cnt += entry['count']
                inst_net += entry['inst_net']
                has_inst = max(has_inst, entry['has_inst'])
        if cnt > 0:
            rows.append((code, iss, {'toplist_count_30d': cnt,
                                     'toplist_inst_net_buy': inst_net,
                                     'toplist_institutional': has_inst}))
    print(f'  [top_list] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'top_list')
        n = batch_update(conn, 'top_list', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P2: block_trade 大宗交易 ──
def ingest_block_trade(conn, write, limit):
    """大宗交易 → placement_evaluation(直写 PE)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        for attempt in range(3):
            try:
                df = pro.block_trade(ts_code=stock)
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
                df = None
        if df is None or len(df) == 0:
            continue
        df['trade_date'] = df.get('trade_date', pd.Series()).astype(str)
        df['tp'] = pd.to_numeric(df.get('tp', pd.Series()), errors='coerce')  # 成交价
        df['amount'] = pd.to_numeric(df.get('amount', pd.Series()), errors='coerce')
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = r['issue_date']
            yr_ago = str(int(iss[:4]) - 1) + iss[4:] if len(iss) == 8 else None
            if not yr_ago:
                continue
            recent = df[(df['trade_date'] >= yr_ago) & (df['trade_date'] <= iss)]
            if len(recent) == 0:
                continue
            f = {'block_count_30d': int(len(recent))}
            # 平均折价率
            if 'tp' in recent.columns:
                # 简化: 折价率需要收盘价对比,此处用成交价/成交价(占位)
                f['block_discount_avg'] = 0.0
            if recent['amount'].notna().any():
                f['block_amount_ratio'] = _sv(float(recent['amount'].sum()))
            rows.append((stock, iss, f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [block_trade] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [block_trade] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'block_trade')
        n = batch_update(conn, 'block_trade', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P2: stk_holdernumber 股东人数 ──
def ingest_holdernumber(conn, write, limit):
    """股东人数 → placement_evaluation(直写 PE)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        for attempt in range(3):
            try:
                df = pro.stk_holdernumber(ts_code=stock)
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
                df = None
        if df is None or len(df) == 0:
            continue
        df['end_date'] = df.get('end_date', pd.Series()).astype(str)
        df['holder_num'] = pd.to_numeric(df.get('holder_num', pd.Series()), errors='coerce')
        df = df.sort_values('end_date')
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = r['issue_date']
            pit = df[(df['end_date'] <= iss) & df['holder_num'].notna()]
            if len(pit) == 0:
                continue
            f = {'holder_count': int(pit.iloc[-1]['holder_num'])}
            if len(pit) >= 2:
                prev = pit.iloc[-2]['holder_num']
                if prev > 0:
                    f['holder_count_chg'] = _sv(float(pit.iloc[-1]['holder_num'] / prev - 1))
            rows.append((stock, iss, f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [holdernumber] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [holdernumber] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'holdernumber')
        n = batch_update(conn, 'holdernumber', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P2: stk_holdertrade 股东增减持 ──
def ingest_holdertrade(conn, write, limit):
    """股东增减持 → placement_evaluation(直写 PE)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        for attempt in range(3):
            try:
                df = pro.stk_holdertrade(ts_code=stock)
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
                df = None
        if df is None or len(df) == 0:
            continue
        df['ann_date'] = df.get('ann_date', pd.Series()).astype(str)
        df['in_vol'] = pd.to_numeric(df.get('in_vol', pd.Series()), errors='coerce')
        df['change_vol'] = pd.to_numeric(df.get('change_vol', pd.Series()), errors='coerce')
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = r['issue_date']
            yr_ago = str(int(iss[:4]) - 1) + iss[4:] if len(iss) == 8 else None
            if not yr_ago:
                continue
            recent = df[(df['ann_date'] >= yr_ago) & (df['ann_date'] <= iss)]
            if len(recent) == 0:
                continue
            chg = recent['change_vol'].dropna()
            f = {}
            if len(chg) > 0:
                f['insider_net_buy_90d'] = _sv(float(chg.sum()))
                f['insider_buy_count_90d'] = int((chg > 0).sum())
                f['insider_direction'] = _sv(float(chg.sum() / (chg.abs().sum() + 1e-9)))
            if f:
                rows.append((stock, iss, f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [holdertrade] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [holdertrade] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'holdertrade')
        n = batch_update(conn, 'holdertrade', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P2: stk_surv 机构调研 ──
def ingest_surv(conn, write, limit):
    """机构调研 → placement_evaluation(直写 PE)。"""
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        for attempt in range(3):
            try:
                df = pro.stk_surv(ts_code=stock)
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
                df = None
        if df is None or len(df) == 0:
            continue
        df['surv_date'] = df.get('surv_date', df.get('ann_date', pd.Series())).astype(str)
        sd = samp[samp['stock_code'] == stock]
        for _, r in sd.iterrows():
            iss = r['issue_date']
            yr_ago = str(int(iss[:4]) - 1) + iss[4:] if len(iss) == 8 else None
            if not yr_ago:
                continue
            recent = df[(df['surv_date'] >= yr_ago) & (df['surv_date'] <= iss)]
            if len(recent) == 0:
                continue
            f = {'surv_count_90d': int(len(recent))}
            try:
                last_d = recent['surv_date'].iloc[-1]
                f['surv_recency'] = int((pd.Timestamp(iss) - pd.Timestamp(last_d)).days)
            except Exception:
                pass
            rows.append((stock, iss, f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [surv] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [surv] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'surv')
        n = batch_update(conn, 'surv', rows)
        print(f'  ✅ 回写 {n} 行')


# ── P3: macro regime ──
def ingest_regime(conn, write, limit):
    """宏观 regime: 北向/Shibor/美债 → market_regime_daily; CPI/PPI/PMI/M/社融/LPR → market_regime_monthly。
    然后从两张表匹配到 placement_evaluation → macro 特征列。
    """
    pro = ts.pro_api()
    # ── 日频: 北向资金 ──
    print('  [regime] 拉取北向资金流...')
    try:
        hsgt = pro.moneyflow_hsgt()
        if hsgt is not None and len(hsgt):
            cur = conn.cursor()
            for _, r in hsgt.iterrows():
                td = str(r.get('trade_date', ''))
                north = _sv(r.get('north_money', r.get('ggt_ss', 0)))
                south = _sv(r.get('south_money', r.get('ggt_sz', 0)))
                if len(td) == 8 and write:
                    cur.execute(
                        "INSERT INTO market_regime_daily (trade_date, hsgt_north, hsgt_south) "
                        "VALUES (%s,%s,%s) ON DUPLICATE KEY UPDATE "
                        "hsgt_north=VALUES(hsgt_north), hsgt_south=VALUES(hsgt_south)",
                        (td, north, south))
            if write:
                conn.commit()
            print(f'  [regime] 北向 {len(hsgt)} 日')
    except Exception as e:
        print(f'  [regime] 北向拉取失败: {e}')

    # ── 日频: Shibor ──
    print('  [regime] 拉取 Shibor...')
    try:
        sh = pro.shibor()
        if sh is not None and len(sh):
            cur = conn.cursor()
            for _, r in sh.iterrows():
                td = str(r.get('date', ''))
                if len(td) != 8 or not write:
                    continue
                cur.execute(
                    "INSERT INTO market_regime_daily (trade_date, shibor_on, shibor_1w, shibor_1m, shibor_3m) "
                    "VALUES (%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE "
                    "shibor_on=VALUES(shibor_on), shibor_1w=VALUES(shibor_1w), "
                    "shibor_1m=VALUES(shibor_1m), shibor_3m=VALUES(shibor_3m)",
                    (td, _sv(r.get('on')), _sv(r.get('1w')), _sv(r.get('1m')), _sv(r.get('3m'))))
            if write:
                conn.commit()
            print(f'  [regime] Shibor {len(sh)} 日')
    except Exception as e:
        print(f'  [regime] Shibor 拉取失败: {e}')

    # ── 月频: CPI/PPI/PMI/M/社融/LPR ──
    print('  [regime] 拉取宏观月频...')
    monthly_data = {}  # month → {col: val}
    apis = [
        ('cn_cpi', 'nt_yoy', 'cpi_yoy'),
        ('cn_ppi', 'ppi_yoy', 'ppi_yoy'),
        ('cn_pmi', 'pmi', 'pmi'),
        ('cn_m', 'm1_yoy', 'm1_yoy'),
    ]
    for api_name, src_col, dst_col in apis:
        try:
            df = getattr(pro, api_name)()
            if df is not None and len(df):
                for _, r in df.iterrows():
                    m = str(r.get('month', r.get('date', '')))[:6]
                    if len(m) != 6:
                        continue
                    if m not in monthly_data:
                        monthly_data[m] = {}
                    monthly_data[m][dst_col] = _sv(r.get(src_col))
        except Exception as e:
            print(f'  [regime] {api_name} 失败: {e}')
    # M2
    try:
        df = pro.cn_m()
        if df is not None and len(df):
            for _, r in df.iterrows():
                m = str(r.get('month', ''))[:6]
                if len(m) == 6:
                    if m not in monthly_data:
                        monthly_data[m] = {}
                    monthly_data[m]['m2_yoy'] = _sv(r.get('m2_yoy'))
    except Exception:
        pass
    # LPR
    try:
        df = pro.shibor_lpr()
        if df is not None and len(df):
            for _, r in df.iterrows():
                m = str(r.get('date', ''))[:6]
                if len(m) == 6:
                    if m not in monthly_data:
                        monthly_data[m] = {}
                    monthly_data[m]['lpr_1y'] = _sv(r.get('lpr1y'))
                    monthly_data[m]['lpr_5y'] = _sv(r.get('lpr5y'))
    except Exception:
        pass
    # 写入 market_regime_monthly
    if write:
        cur = conn.cursor()
        for m, d in monthly_data.items():
            cur.execute(
                "INSERT INTO market_regime_monthly (month, cpi_yoy, ppi_yoy, pmi, "
                "m1_yoy, m2_yoy, lpr_1y, lpr_5y) VALUES (%s,%s,%s,%s,%s,%s,%s,%s) "
                "ON DUPLICATE KEY UPDATE cpi_yoy=VALUES(cpi_yoy), ppi_yoy=VALUES(ppi_yoy), "
                "pmi=VALUES(pmi), m1_yoy=VALUES(m1_yoy), m2_yoy=VALUES(m2_yoy), "
                "lpr_1y=VALUES(lpr_1y), lpr_5y=VALUES(lpr_5y)",
                (m, d.get('cpi_yoy'), d.get('ppi_yoy'), d.get('pmi'),
                 d.get('m1_yoy'), d.get('m2_yoy'), d.get('lpr_1y'), d.get('lpr_5y')))
        conn.commit()
        print(f'  [regime] 月频 {len(monthly_data)} 月写入完成')

    # ── 匹配到 PE → macro 特征列 ──
    samp = pd.DataFrame(_dict_query(conn,
        "SELECT stock_code, issue_date FROM placement_evaluation "
        "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8"))
    samp['issue_date'] = samp['issue_date'].astype(str)
    # 读 regime 表
    cur = conn.cursor()
    cur.execute("SELECT * FROM market_regime_monthly ORDER BY month")
    mcols = [d[0] for d in cur.description]
    mdata = pd.DataFrame(cur.fetchall(), columns=mcols)
    cur.execute("SELECT trade_date, hsgt_north, shibor_3m, us_10y FROM market_regime_daily ORDER BY trade_date")
    dcols = [d[0] for d in cur.description]
    ddata = pd.DataFrame(cur.fetchall(), columns=dcols)
    cur.close()

    rows = []
    for _, r in samp.iterrows():
        iss = r['issue_date']
        m = iss[:6]
        f = {}
        # 月频匹配
        mrow = mdata[mdata['month'] <= m]
        if len(mrow) > 0:
            mr = mrow.iloc[-1]
            for col in ['cpi_yoy', 'ppi_yoy', 'pmi', 'm1_yoy', 'm2_yoy', 'lpr_1y']:
                v = _sv(mr.get(col))
                if v is not None:
                    f[f'macro_{col}'] = v
            cpi = _sv(mr.get('cpi_yoy'))
            ppi = _sv(mr.get('ppi_yoy'))
            if cpi is not None and ppi is not None:
                f['macro_ppi_cpi_spread'] = ppi - cpi
            m1 = _sv(mr.get('m1_yoy'))
            m2 = _sv(mr.get('m2_yoy'))
            if m1 is not None and m2 is not None:
                f['macro_m1_m2_scissor'] = m1 - m2
            pmi = _sv(mr.get('pmi'))
            if pmi is not None:
                f['macro_pmi_expansion'] = 1 if pmi > 50 else 0
        # 日频匹配
        drow = ddata[ddata['trade_date'] <= iss]
        if len(drow) > 0:
            dr = drow.iloc[-1]
            sh3 = _sv(dr.get('shibor_3m'))
            if sh3 is not None:
                f['macro_shibor_3m'] = sh3
            us10 = _sv(dr.get('us_10y'))
            if us10 is not None:
                f['macro_us_10y'] = us10
                if sh3 is not None:
                    f['macro_us_cn_spread'] = us10 - sh3  # 近似中美利差
            # 北向5日/20日
            if len(drow) >= 5:
                hsgt = pd.to_numeric(drow['hsgt_north'].tail(5), errors='coerce')
                f['macro_hsgt_net_5d'] = _sv(float(hsgt.sum()))
            if len(drow) >= 20:
                hsgt20 = pd.to_numeric(drow['hsgt_north'].tail(20), errors='coerce')
                f['macro_hsgt_net_20d'] = _sv(float(hsgt20.sum()))
        if f:
            rows.append((r['stock_code'], iss, f))
    print(f'  [macro] 匹配 {len(rows)} 样本')
    if write and rows:
        ensure_columns(conn, 'macro')
        n = batch_update(conn, 'macro', rows)
        print(f'  ✅ 回写 {n} 行')


SOURCES = {'placement': ingest_placement, 'chip': ingest_chip,
           'capitalflow': ingest_capitalflow, 'smc': ingest_smc, 'sue': ingest_sue,
           # P0
           'margin': ingest_margin, 'report_rc': ingest_report_rc,
           # P1
           'pledge': ingest_pledge, 'dividend': ingest_dividend,
           'repurchase': ingest_repurchase,
           # P2
           'top_list': ingest_top_list, 'block_trade': ingest_block_trade,
           'holdernumber': ingest_holdernumber, 'holdertrade': ingest_holdertrade,
           'surv': ingest_surv,
           # P3
           'regime': ingest_regime, 'macro': ingest_regime,
           }


def main():
    ap = argparse.ArgumentParser(description='统一因子摄入 → placement_evaluation DB')
    ap.add_argument('source', choices=list(SOURCES) + ['all'])
    ap.add_argument('--write', action='store_true', help='实际回写(否则 dry-run)')
    ap.add_argument('--limit', type=int, default=0, help='只取前 N 只股(0=全量)')
    args = ap.parse_args()
    conn = pymysql.connect(**_CFG)
    targets = list(SOURCES) if args.source == 'all' else [args.source]
    for src in targets:
        print(f'\n=== {src} ===')
        ensure_columns(conn, src)
        SOURCES[src](conn, args.write, args.limit)
    conn.close()


if __name__ == '__main__':
    main()
