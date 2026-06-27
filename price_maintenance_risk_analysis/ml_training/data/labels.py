#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""labels.py — 标签计算(合并 compute_labels + recompute_label_qfq + backfill_evaluations)。
子命令: python labels.py <labels|qfq|evaluations> [...]

吸收原 2 脚本: excess(超额收益 excess_mkt/ind_{1,3,7}m) / shortterm(短线 return_1w/2w/4w)。
**只写原始值到 DB**, 不写标签列到 parquet——阈值/极性标签由 export_features 统一发射(与 标签_盈利_* 一致)。
return_*m 由 recompute_label_qfq.py(既有)负责, 本脚本不管。

用法:
  python scripts/compute_labels.py {excess|shortterm|all}
"""
import argparse
import bisect
import calendar
import os
import sys
import time
from datetime import datetime

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
EXCESS_COLS = [f'excess_{b}_{h}m' for b in ('mkt', 'ind') for h in (1, 3, 7)]
SHORT_COLS = ['return_1w', 'return_2w', 'return_4w']
WINDOWS = {'1w': 5, '2w': 10, '4w': 20}   # 交易日



# ====== compute_labels (excess/shortterm 标签) ======
def _clean(v):
    try:
        f = float(v)
        return None if f != f else f
    except (TypeError, ValueError):
        return None


def add_months(ymd, months):
    y, m, d = int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8])
    tot = y * 12 + (m - 1) + months
    ny, nm = tot // 12, tot % 12 + 1
    last = calendar.monthrange(ny, nm)[1]
    return f'{ny:04d}{nm:02d}{min(d, last):02d}'


def build_series(close_map):
    arr = []
    for k, v in close_map.items():
        try:
            f = float(v)
        except (TypeError, ValueError):
            continue
        if f != f:
            continue
        try:
            arr.append((datetime.strptime(k, '%Y%m%d').toordinal(), f))
        except ValueError:
            continue
    arr.sort()
    return arr


def _nearest(series, target_ord, tol=15):
    if not series:
        return None
    i = bisect.bisect_left(series, (target_ord, -1e18))
    cands = [series[j] for j in (i - 1, i) if 0 <= j < len(series)]
    if not cands:
        return None
    o, c = min(cands, key=lambda x: abs(x[0] - target_ord))
    return (o, c) if abs(o - target_ord) <= tol else None


def bench_return(series, issue, h):
    try:
        t0 = datetime.strptime(issue, '%Y%m%d').toordinal()
        t1 = datetime.strptime(add_months(issue, h), '%Y%m%d').toordinal()
    except ValueError:
        return None
    c0 = _nearest(series, t0); c1 = _nearest(series, t1)
    if c0 is None or c1 is None or c0[1] == 0:
        return None
    return (c1[1] / c0[1] - 1) * 100


# ── excess ──
def ingest_excess(conn):
    cur = conn.cursor(cursorclass=pymysql.cursors.DictCursor)
    cur.execute("SELECT stock_code, issue_date, return_1m, return_3m, return_7m "
                "FROM placement_evaluation WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8")
    samp = pd.DataFrame(cur.fetchall())
    conn.cursor().close()
    samp['issue_date'] = samp['issue_date'].astype(str)
    print(f'  [excess] 样本 {len(samp)}')
    mkt = ts.pro_api().index_daily(ts_code='000300.SH').sort_values('trade_date')
    mkt_series = build_series(dict(zip(mkt['trade_date'].astype(str), pd.to_numeric(mkt['close'], errors='coerce'))))
    conn_plain = pymysql.connect(**_CFG)
    ind = pd.read_sql('SELECT index_code, trade_date, close FROM industry_daily', conn_plain)
    ind_series = {code: build_series(dict(zip(g['trade_date'].astype(str), pd.to_numeric(g['close'], errors='coerce'))))
                  for code, g in ind.groupby('index_code')}
    smap = pd.read_sql('SELECT stock_code, index_code FROM industry_data', conn_plain)
    conn_plain.close()
    stock2ind = dict(zip(smap['stock_code'], smap['index_code']))
    print(f'  [excess] 大盘 {len(mkt_series)}日 | 行业 {len(ind_series)}指数')

    set_c = ', '.join(f'{c}=%s' for c in EXCESS_COLS)
    upd = f'UPDATE placement_evaluation SET {set_c} WHERE stock_code=%s AND issue_date=%s'
    vals = []
    for _, r in samp.iterrows():
        iss = r['issue_date']; out = {}
        for h in (1, 3, 7):
            sret = r.get(f'return_{h}m')
            if pd.isna(sret):
                continue
            sret = float(sret)
            mr = bench_return(mkt_series, iss, h)
            if mr is not None:
                out[f'excess_mkt_{h}m'] = sret - mr
            iser = ind_series.get(stock2ind.get(r['stock_code']))
            if iser is not None:
                ir = bench_return(iser, iss, h)
                if ir is not None:
                    out[f'excess_ind_{h}m'] = sret - ir
        vals.append(tuple(_clean(out.get(c)) for c in EXCESS_COLS) + (r['stock_code'], iss))
    n = conn.cursor().executemany(upd, vals)
    conn.commit()
    print(f'  ✅ 回写 {n} 行 excess')


# ── shortterm ──
def ingest_shortterm(conn, limit=0):
    cur = conn.cursor(cursorclass=pymysql.cursors.DictCursor)
    cur.execute("SELECT stock_code, issue_date FROM placement_evaluation "
                "WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8")
    samp = pd.DataFrame(cur.fetchall())
    conn.cursor().close()
    samp['issue_date'] = samp['issue_date'].astype(str)
    stocks = sorted(samp['stock_code'].unique())
    if limit:
        stocks = stocks[:limit]
    pro = ts.pro_api()
    set_c = ', '.join(f'{c}=%s' for c in SHORT_COLS)
    upd = f'UPDATE placement_evaluation SET {set_c} WHERE stock_code=%s AND issue_date=%s'
    vals = []; done = 0
    for si, stock in enumerate(stocks):
        grp = samp[samp['stock_code'] == stock]
        dates = sorted(grp['issue_date'].tolist())
        sd = min(dates)[:4] + '0101'; ed = max(dates)[:4] + '1231'
        cmap = {}
        for attempt in range(3):
            try:
                d = ts.pro_bar(ts_code=stock, start_date=sd, end_date=ed, adj='qfq')
                if d is not None and len(d):
                    d = d.sort_values('trade_date')
                    cmap = dict(zip(d['trade_date'].astype(str), pd.to_numeric(d['close'], errors='coerce')))
                break
            except Exception:
                time.sleep(1.0 * (attempt + 1))
        time.sleep(0.3)
        if not cmap:
            continue
        tds = sorted(cmap.keys())
        for _, r in grp.iterrows():
            iss = r['issue_date']
            i0 = bisect.bisect_left(tds, iss)
            if i0 == len(tds):
                continue
            if tds[i0] > iss and i0 > 0:
                i0 -= 1
            c0 = cmap.get(tds[i0])
            if c0 is None or c0 == 0:
                continue
            out = {}
            for tag, n in WINDOWS.items():
                j = i0 + n
                if j < len(tds):
                    c1 = cmap.get(tds[j])
                    if c1 is not None:
                        out[f'return_{tag}'] = (c1 / c0 - 1) * 100
            vals.append(tuple(_clean(out.get(c)) for c in SHORT_COLS) + (stock, iss))
        if (si + 1) % 200 == 0:
            print(f'  [shortterm] {si+1}/{len(stocks)} | {len(vals)} 样本', flush=True)
    n = conn.cursor().executemany(upd, vals)
    conn.commit()
    print(f'  ✅ 回写 {n} 行 shortterm')


SOURCES = {'excess': ingest_excess, 'shortterm': ingest_shortterm}


def ensure_columns(conn, source):
    cols = EXCESS_COLS if source == 'excess' else SHORT_COLS
    cur = conn.cursor()
    cur.execute("""SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
                   WHERE TABLE_SCHEMA=%s AND TABLE_NAME='placement_evaluation'""", (_CFG['database'],))
    have = {r[0] for r in cur.fetchall()}
    miss = [c for c in cols if c not in have]
    if miss:
        cur.execute('ALTER TABLE placement_evaluation ADD COLUMN ' + ', ADD COLUMN '.join(f'{c} DOUBLE' for c in miss))
        conn.commit()
        print(f'  [{source}] 补列: {miss}')


def main_labels():
    ap = argparse.ArgumentParser(description='统一标签原始值 → placement_evaluation DB')
    ap.add_argument('source', choices=list(SOURCES) + ['all'])
    ap.add_argument('--limit', type=int, default=0)
    args = ap.parse_args()
    conn = pymysql.connect(**_CFG)
    targets = list(SOURCES) if args.source == 'all' else [args.source]
    for src in targets:
        print(f'\n=== {src} ===')
        ensure_columns(conn, src)
        if src == 'shortterm':
            ingest_shortterm(conn, args.limit)
        else:
            SOURCES[src](conn)
    conn.close()

# ====== recompute_label_qfq (qfq 重算) ======
def _add_months(ymd, months):
    """YYYYMMDD 字符串 +N 月 → YYYYMMDD 字符串(粗略, 日溢出截到月末)。"""
    y, m, d = int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8])
    tot = y * 12 + (m - 1) + months
    ny, nm = tot // 12, tot % 12 + 1
    import calendar
    last = calendar.monthrange(ny, nm)[1]
    return f'{ny:04d}{nm:02d}{min(d, last):02d}'

def _nearest_close(close_map, target, tol_days=10):
    """close_map: {trade_date_str: close}; 找离 target 日期最近(<=tol_days)的收盘。"""
    try:
        t = datetime.strptime(target, '%Y%m%d')
    except ValueError:
        return None
    best_k, best_dd = None, 1e9
    for k, v in close_map.items():
        if v != v:  # NaN
            continue
        try:
            kd = datetime.strptime(k, '%Y%m%d')
        except ValueError:
            continue
        dd = abs((kd - t).days)
        if dd < best_dd:
            best_dd, best_k = dd, k
    if best_k is None or best_dd > tol_days:
        return None
    return float(close_map[best_k])

def _ensure_horizon_columns(conn, horizons):
    """幂等: 缺失的 return_{h}m/price_{h}m 列补上(自愈, 兼容旧 MySQL 无 IF NOT EXISTS)。"""
    need = [f'return_{h}m' for h in horizons] + [f'price_{h}m' for h in horizons]
    with conn.cursor() as cur:
        cur.execute(
            "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
            "WHERE TABLE_SCHEMA=%s AND TABLE_NAME='placement_evaluation'", (_CFG['database'],))
        have = {r[0] for r in cur.fetchall()}
    missing = [c for c in need if c not in have]
    if missing:
        ddl = 'ALTER TABLE placement_evaluation ' + ', '.join(f'ADD COLUMN {c} DOUBLE' for c in missing)
        with conn.cursor() as cur:
            cur.execute(ddl)
        conn.commit()
        print(f'  已补列: {missing}')
def main_qfq():
    ap = argparse.ArgumentParser()
    ap.add_argument('--horizons', default='1,3,6,7,12',
                        help='期限(月), 逗号分隔, 默认 1,3,6,7,12')
    ap.add_argument('--dry-run', action='store_true', help='只打印不回写')
    args = ap.parse_args()
    horizons = sorted({int(x) for x in args.horizons.split(',') if x.strip()})
    max_h = max(horizons)
    conn = pymysql.connect(**_CFG)
    _ensure_horizon_columns(conn, horizons)
    samples = pd.read_sql(
            "SELECT stock_code, issue_date FROM placement_evaluation "
            "WHERE issue_date IS NOT NULL AND issue_date<>'' AND LENGTH(issue_date)=8", conn)
    samples['issue_date'] = samples['issue_date'].astype(str)
    print(f'样本 {len(samples)} 条 / {samples["stock_code"].nunique()} 只股票 | 期限 {horizons}')
    set_clause = ', '.join([f'return_{h}m=%s, price_{h}m=%s' for h in horizons])
    upd_sql = (f'UPDATE placement_evaluation SET {set_clause} '
                   'WHERE stock_code=%s AND issue_date=%s')
    n_slots = 2 * len(horizons)
    pro = ts.pro_api(os.environ['TUSHARE_TOKEN'])
    cur = conn.cursor()
    ok = upd = fail = 0
    for stock, grp in samples.groupby('stock_code'):
            stock = str(stock)
            # 取数范围覆盖 issue 与最大期限
            dates = []
            for _, r in grp.iterrows():
                dates.append(r['issue_date'])
                dates.append(_add_months(r['issue_date'], max_h))
            sd = (min(dates)[:4] + '0101'); ed = (max(dates)[:4] + '1231')
            try:
                df = None
                for attempt in range(3):
                    try:
                        df = ts.pro_bar(ts_code=stock, start_date=sd, end_date=ed, adj='qfq')
                    except Exception:
                        df = None
                    if df is not None and len(df) > 0:
                        break
                    time.sleep(1.2 * (attempt + 1))
                time.sleep(0.3)
                if df is None or len(df) == 0:
                    fail += 1; continue
                cmap = dict(zip(df['trade_date'].astype(str),
                                pd.to_numeric(df['close'], errors='coerce')))
            except Exception as e:
                fail += 1; print(f'  {stock} 取数失败: {e}'); continue

            for _, r in grp.iterrows():
                issue = r['issue_date']
                c_i = _nearest_close(cmap, issue)
                if c_i is None or c_i == 0:
                    continue
                params = []
                any_h = False
                for h in horizons:
                    c_t = _nearest_close(cmap, _add_months(issue, h))
                    if c_t is None:
                        ret = None
                    else:
                        ret = round((c_t - c_i) / c_i * 100, 4)
                        any_h = True
                    params += [ret, round(c_t, 4) if c_t is not None else None]
                params += [stock, issue]
                if not any_h:
                    continue
                ok += 1
                if not args.dry_run:
                    assert len(params) == n_slots + 2
                    cur.execute(upd_sql, params)
                    upd += 1
            if not args.dry_run:
                conn.commit()
            done = ok + fail
            if done and done % 50 == 0:
                print(f'  进度: ok={ok}, updated={upd}, failed_stocks={fail}', flush=True)
    if not args.dry_run:
            conn.commit()
    conn.close()
    print(f'完成: {ok} 样本重算, {upd} 行回写, {fail} 股取数失败 | 期限 {horizons}')
_PROJ_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SUB_SCENARIO_MAP = {
    '市场指数': 'sub_market_index',
    '行业PE': 'sub_industry_pe',
    '个股PE': 'sub_stock_pe',
    'DCF估值': 'sub_dcf',
    '修正PE估值': 'sub_adj_pe',
    '参数构造': 'sub_param_build',
    '蒙特卡洛': 'sub_monte_carlo',
    '反向推算': 'sub_reverse_calc',
}
RELATIVE_LABELS = ['T-4', 'T-3', 'T-2', 'T-1', 'T']

# ====== backfill_evaluations ======
def _to_float(v):
    """安全转 float；空/'-' → None。不剥离 %。"""
    if v is None:
        return None
    s = str(v).strip()
    if s == '' or s == '-' or s.lower() == 'nan':
        return None
    try:
        return float(s.replace(',', ''))
    except (ValueError, TypeError):
        return None

def _pct_to_decimal(v):
    """百分比字符串 → 小数: '-20.00%' → -0.2；'-' → None。"""
    if v is None:
        return None
    s = str(v).strip().replace('%', '').replace('+', '').replace(' ', '')
    if s == '' or s == '-' or s.lower() == 'nan':
        return None
    try:
        return float(s) / 100.0
    except (ValueError, TypeError):
        return None

def _ret_to_pct(v):
    """涨跌幅百分比字符串 → 数值: '-24.91%' → -24.91（保留百分比数值，非小数）。"""
    if v is None:
        return None
    s = str(v).strip().replace('%', '').replace('+', '').replace(' ', '')
    if s == '' or s == '-' or s.lower() == 'nan':
        return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None

def _issue_date_str(v):
    """报价日(数字 20200109.0 或 '20200109') → '20200109' 字符串。"""
    if v is None:
        return None
    try:
        return str(int(float(v)))
    except (ValueError, TypeError):
        s = str(v).strip().replace('-', '').replace('/', '')
        return s if s else None

def _sub_flag(v):
    """子场景值: 含'✓' → 1，否则 → 0。"""
    return 1 if (v is not None and '✓' in str(v)) else 0

def _row_to_placement_eval(row, batch_id):
    """单行 → placement_evaluation dict（供 save_placement_evaluation）。"""
    return {
        'stock_code': str(row.get('股票代码', '')).strip(),
        'stock_name': str(row.get('股票简称', '') or ''),
        'batch_id': batch_id,
        'issue_date': _issue_date_str(row.get('报价日')),
        'issue_date_price': _to_float(row.get('报价日价格')),
        'valid_thresholds': int(_to_float(row.get('有效阈值数')) or 0) if _to_float(row.get('有效阈值数')) is not None else None,
        'premium_min': _pct_to_decimal(row.get('溢价率下限')),
        'premium_max': _pct_to_decimal(row.get('溢价率上限')),
        'decision': str(row.get('定增决策', '') or ''),
        # 子场景
        **{db_col: _sub_flag(row.get(excel_col)) for excel_col, db_col in SUB_SCENARIO_MAP.items()},
        # 行业
        'industry_l1': str(row.get('一级行业', '') or ''),
        'industry_l2': str(row.get('二级行业', '') or ''),
        'industry_l3': str(row.get('三级行业', '') or ''),
        # 趋势
        'total_slope': _to_float(row.get('总分_斜率')),
        'total_trend': str(row.get('总分_趋势', '') or ''),
        'profit_slope': _to_float(row.get('盈利能力_斜率')),
        'profit_trend': str(row.get('盈利能力_趋势', '') or ''),
        'growth_slope': _to_float(row.get('成长能力_斜率')),
        'growth_trend': str(row.get('成长能力_趋势', '') or ''),
        'combined_trend': str(row.get('综合趋势', '') or ''),
        # 标签
        'return_7m': _ret_to_pct(row.get('7个月后涨跌幅')),
        'price_7m': _to_float(row.get('7个月后价格')),
        'final_conclusion': str(row.get('最终结论', '') or ''),
    }

def _row_to_annual_scores(row):
    """单行 → {report_year: score_dict}（供 save_annual_scores）。

    评分列用相对年份 T-4..T，按「报价日」年份回溯成绝对年份:
        T-4 → base_year-4, ..., T → base_year
    （与 export_features.load_scored_features_from_db 的反查逻辑一致）
    """
    issue_date = _issue_date_str(row.get('报价日'))
    if not issue_date or len(issue_date) < 4:
        return {}
    try:
        base_year = int(issue_date[:4])
    except ValueError:
        return {}

    scores_by_year = {}
    industry = {
        'industry_l1': str(row.get('一级行业', '') or ''),
        'industry_l2': str(row.get('二级行业', '') or ''),
        'industry_l3': str(row.get('三级行业', '') or ''),
    }
    stock_name = str(row.get('股票简称', '') or '')

    for i, label in enumerate(RELATIVE_LABELS):
        year = base_year - 4 + i
        total = _to_float(row.get(f'总分_{label}'))
        # 该行该年若无评分则跳过
        if total is None:
            continue
        scores_by_year[year] = {
            'stock_name': stock_name,
            'total_score': total,
            'rating': str(row.get(f'评级_{label}', '') or ''),
            'profitability': _to_float(row.get(f'盈利能力_{label}')),
            'growth': _to_float(row.get(f'成长能力_{label}')),
            'operating': None,   # Excel 无此维度
            'solvency': None,    # Excel 无此维度
            **industry,
        }
    return scores_by_year

def backfill_one_file(db, path, batch_id):
    """回填单个 Excel 文件，返回 (n_eval, n_scores, n_skipped)。"""
    if not os.path.exists(path):
        print(f'  ✗ 文件不存在: {path}')
        return 0, 0, 0

    df = pd.read_excel(path, sheet_name=0)
    print(f'\n  📄 {os.path.basename(path)}: {len(df)} 行 × {len(df.columns)} 列')
    print(f'     报价日 非空: {df["报价日"].notna().sum() if "报价日" in df.columns else 0}'
          f' | 7个月后涨跌幅 非空: {df["7个月后涨跌幅"].notna().sum() if "7个月后涨跌幅" in df.columns else 0}')

    n_eval = n_scores = n_skipped = 0
    for _, row in df.iterrows():
        code = str(row.get('股票代码', '')).strip()
        if not code:
            n_skipped += 1
            continue

        # 1. placement_evaluation
        pe = _row_to_placement_eval(row, batch_id)
        if pe['issue_date']:
            try:
                db.save_placement_evaluation(pe)
                n_eval += 1
            except Exception as e:
                print(f'     ⚠ {code} placement 落库失败: {e}')

        # 2. company_annual_scores
        scores_by_year = _row_to_annual_scores(row)
        if scores_by_year:
            try:
                db.save_annual_scores(code, scores_by_year, batch_id=batch_id)
                n_scores += len(scores_by_year)
            except Exception as e:
                print(f'     ⚠ {code} annual 落库失败: {e}')

    print(f'     ✅ placement_evaluation: {n_eval} 条 | company_annual_scores: {n_scores} 条 | 跳过: {n_skipped}')
    return n_eval, n_scores, n_skipped
def main_evaluations():
    parser = argparse.ArgumentParser(description='历史定增评估数据回填 → placement_evaluation + company_annual_scores')
    parser.add_argument('--labeled', nargs='+', required=True,
                            help='带标签的 scored Excel 路径（可多个）')
    parser.add_argument('--batch-id', default=None,
                            help='批次ID标记（默认 backfill_YYYYMMDD）')
    args = parser.parse_args()
    batch_id = args.batch_id or 'backfill_hist'
    db = ValuationDB()
    print(f'回填目标: investment_valuation.{ "(placement_evaluation, company_annual_scores)" }')
    print(f'批次ID: {batch_id}')
    print(f'待回填文件: {len(args.labeled)} 个')
    total_eval = total_scores = 0
    for path in args.labeled:
            path = os.path.expanduser(path)
            n_eval, n_scores, _ = backfill_one_file(db, path, batch_id)
            total_eval += n_eval
            total_scores += n_scores
    print('\n' + '=' * 60)
    print(f'🎉 回填完成: placement_evaluation 累计写入 {total_eval} 条, '
              f'company_annual_scores 累计写入 {total_scores} 条')
    print('   （均为 upsert，重复行已更新而非重复插入）')
    print('=' * 60)

if __name__ == '__main__':
    _cmd = sys.argv[1] if len(sys.argv) > 1 else 'labels'
    sys.argv = [sys.argv[0]] + sys.argv[2:]
    {'labels': main_labels, 'qfq': main_qfq, 'evaluations': main_evaluations}[_cmd]()
