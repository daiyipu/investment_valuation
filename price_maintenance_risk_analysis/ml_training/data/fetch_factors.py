#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""统一因子摄入 → placement_evaluation DB(单一原始源)。

吸收原 4 个脚本: placement(定增结构) / chip(筹码 cyq_chips) / capitalflow(资金流+北向) / smc(聪明钱 OHLCV)。
全部写 placement_evaluation(key=stock_code+issue_date, 幂等 ADD COLUMN)。SMC 由原"写 parquet"改为写 DB。
评估(IV/AUC)不在此脚本, 归建模侧。

用法:
  python scripts/fetch_factors.py {placement|chip|capitalflow|smc|all} [--write] [--limit N]
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


def ingest_chip(conn, write, limit):
    cur = conn.cursor(cursorclass=pymysql.cursors.DictCursor)
    cur.execute("SELECT stock_code, issue_date, issue_date_price FROM placement_evaluation "
                "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8 AND issue_date >= '2018'")
    samp = pd.DataFrame(cur.fetchall())
    conn.cursor().close()
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        sd = samp[samp['stock_code'] == stock]
        dates = sorted(int(x) for x in sd['issue_date'])
        cm = fetch_stock_chips(pro, stock, dates)
        for _, r in sd.iterrows():
            td = nearest_td(cm, r['issue_date'])
            cp = r.get('issue_date_price')
            f = compute_chip_features(cm.get(td, []), float(cp) if pd.notna(cp) else None) if td else None
            if f:
                rows.append((stock, r['issue_date'], f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [chip] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [chip] 匹配 {len(rows)} 样本')
    if write:
        n = batch_update(conn, 'chip', rows)
        print(f'  ✅ 回写 {n} 行')


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
    cur = conn.cursor(cursorclass=pymysql.cursors.DictCursor)
    cur.execute("SELECT stock_code, issue_date FROM placement_evaluation "
                "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8 AND issue_date >= '2013'")
    samp = pd.DataFrame(cur.fetchall())
    conn.cursor().close()
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    rows = []
    for i, stock in enumerate(stocks):
        sd = samp[samp['stock_code'] == stock]
        dates = sorted(int(x) for x in sd['issue_date'])
        mf = _fetch_series(pro, stock, dates, 'moneyflow')
        nb = _fetch_series(pro, stock, dates, 'hk_hold')
        for _, r in sd.iterrows():
            f = {}
            f.update(mf_features(mf, int(r['issue_date'])))
            f.update(nb_features(nb, int(r['issue_date'])))
            if f:
                rows.append((stock, r['issue_date'], f))
        time.sleep(0.3)
        if (i + 1) % 200 == 0:
            print(f'  [capitalflow] {i+1}/{len(stocks)} | {len(rows)} 样本', flush=True)
    print(f'  [capitalflow] 匹配 {len(rows)} 样本')
    if write:
        n = batch_update(conn, 'capitalflow', rows)
        print(f'  ✅ 回写 {n} 行')


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


SOURCES = {'placement': ingest_placement, 'chip': ingest_chip,
           'capitalflow': ingest_capitalflow, 'smc': ingest_smc, 'sue': ingest_sue}


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
