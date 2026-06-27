# ====== refresh_industry_daily (refresh) ======
import argparse
import os
import sys
import time

import pandas as pd

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # data/→ml_training/→PKG
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'scripts'))
sys.path.insert(0, os.path.join(PKG, 'ml_training'))  # data 兄弟(from data.X)

from tushare_token import resolve_tushare_token  # noqa: E402
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
from data.update_market_data import fetch_industry_index_data  # noqa: E402
from utils.db_manager import ValuationDB  # noqa: E402

import pymysql  # noqa: E402


def _f_mv(x):
    """nan/inf→None(pymysql 拒 nan)。"""
    try:
        x = float(x); return None if x != x or x in (float('inf'), float('-inf')) else x
    except Exception:
        return None


def ingest_idx_factor_pro(indices):
    """刷新 index_factor_pro(指数技术面 idx_factor_pro, 87因子)→ 时序表 index_factor_pro。
    indices = 指数码列表(申万行业 + 大盘 000300.SH)。每指数全历史1次(8000行/次)。
    PIT: trade_date≤回溯日期 切片做特征。与 industry_daily 同批指数, 统一在此脚本摄入(不另开脚本)。
    幂等建表(ON DUP KEY 更新); _f_mv 处理 NaN(见 save-nan-silent-drop-bug)。"""
    import tushare as ts
    pro = ts.pro_api()
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    cols = None; ok = 0; t0 = time.time()
    with conn.cursor() as cur:
        for n, ic in enumerate(indices, 1):
            df = None
            for a in range(3):
                try:
                    df = pro.idx_factor_pro(ts_code=ic); break
                except Exception:
                    time.sleep(1.5 * (a + 1))
            if df is None or len(df) == 0:
                continue
            df = df.sort_values('trade_date')
            factor_cols = [c for c in df.columns if c not in ('ts_code', 'trade_date')]
            if cols is None:
                cols = factor_cols
                coldef = ','.join(f'`{c}` DOUBLE' for c in cols)
                cur.execute(f'CREATE TABLE IF NOT EXISTS index_factor_pro (index_code VARCHAR(16) NOT NULL, '
                            f'trade_date CHAR(8) NOT NULL, {coldef}, PRIMARY KEY(index_code,trade_date)) '
                            f'ENGINE=InnoDB DEFAULT CHARSET=utf8mb4')
                print(f'  idx_factor_pro 建表: {len(cols)} 因子列')
            placeholders = ','.join(['%s'] * (2 + len(cols)))
            collist = 'index_code,trade_date,' + ','.join(f'`{c}`' for c in cols)
            upd = ','.join(f'`{c}`=VALUES(`{c}`)' for c in cols)
            sql = f'INSERT INTO index_factor_pro ({collist}) VALUES ({placeholders}) ON DUPLICATE KEY UPDATE {upd}'
            rows = [tuple([ic, str(r['trade_date'])] + [_f_mv(r[c]) for c in cols]) for _, r in df.iterrows()]
            B = 1000
            for i in range(0, len(rows), B):
                cur.executemany(sql, rows[i:i + B])
            conn.commit(); ok += 1
            if n % 50 == 0:
                print(f'  idx_factor_pro {n}/{len(indices)} (ok {ok}) {time.time()-t0:.0f}s')
    conn.close()
    print(f'✅ idx_factor_pro: {ok}/{len(indices)} 指数, {len(cols) if cols else 0} 因子 ({(time.time()-t0)/60:.1f}min)')


def main_refresh():
    ap = argparse.ArgumentParser(description='按唯一行业指数刷新 industry_daily 全历史(去重+续跑)')
    ap.add_argument('--days', type=int, default=4000, help='回溯天数(默认4000→start≈2004, 覆盖2010+)')
    ap.add_argument('--cutoff', type=int, default=20150101, help='行业 min trade_date≤此值视为已有全历史, 跳过')
    ap.add_argument('--limit', type=int, default=0, help='只处理前 N 个缺早年数据的行业(0=全部, 冒烟用)')
    args = ap.parse_args()

    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    # 唯一行业 + 代表股(取每个 index_code 下的任意一只 stock_code, sw_l3_name, 供 AKShare 降级反查)
    reps = pd.read_sql(
        "SELECT index_code, sw_l3_name, MIN(stock_code) stock_code "
        "FROM industry_data WHERE index_code IS NOT NULL AND index_code<>'' "
        "GROUP BY index_code, sw_l3_name", conn)
    # 已有早年数据的行业(数值比较 trade_date+0, 避免 varchar 词法比较坑): 这些跳过
    done = pd.read_sql(
        f"SELECT DISTINCT index_code FROM industry_daily WHERE trade_date+0 <= {args.cutoff}", conn)
    conn.close()
    done_set = set(done['index_code'].astype(str))

    # 待刷新 = 未在 done_set 里(无早年数据)
    todo = [(str(r['index_code']), r['sw_l3_name'], r['stock_code'])
            for _, r in reps.iterrows() if str(r['index_code']) not in done_set]
    if args.limit:
        todo = todo[:args.limit]

    print(f'唯一行业 {len(reps)} | 已有早年(trade_date≤{args.cutoff})跳过 {len(done_set)} | 待刷新 {len(todo)} (days={args.days})')
    if not todo:
        print('✅ 无待刷新行业'); return

    db = ValuationDB()
    ok = fail = 0
    for i, (ic, name, rep_stock) in enumerate(todo):
        try:
            df = fetch_industry_index_data(ic, days=args.days, stock_code=rep_stock, sw_industry_name=name)
            if df is not None and len(df) >= 250:
                # NaN→None(pymysql 不接受 NaN; save_industry_daily 逐行 row.get 会取到 NaN)
                df = df.replace({float('nan'): None})
                # data_source: tushare_sw 有 pe/pb; AKShare 降级无 pb → 标 akshare_ths
                src = 'tushare_sw'
                pb_col = df.get('pb')
                if pb_col is None or pb_col.dropna().empty:
                    src = 'akshare_ths'
                db.save_industry_daily(ic, df, data_source=src)
                ok += 1
                rng = f"{df['trade_date'].iloc[0]}~{df['trade_date'].iloc[-1]}"
                flag = '' if src == 'tushare_sw' else ' ⚠️AKShare无pb'
                print(f'  [{i+1}/{len(todo)}] {ic} {name}: {len(df)}行 {rng}{flag}')
            else:
                fail += 1
                print(f'  [{i+1}/{len(todo)}] {ic} {name}: 数据不足({0 if df is None else len(df)}行)')
        except Exception as e:
            fail += 1
            print(f'  [{i+1}/{len(todo)}] {ic} {name} 失败: {e}')
        if (i + 1) % 25 == 0:
            print(f'  --- 进度 {i+1}/{len(todo)} (ok={ok} fail={fail})')
        time.sleep(0.3)
    print(f'完成: {ok} 行业刷新成功 / {fail} 失败 / 共 {len(todo)}')

    # 同批指数刷新 index_factor_pro(指数技术面; 大盘+申万行业; 统一在此脚本, 不另开)
    all_idx = sorted({str(r['index_code']) for _, r in reps.iterrows()}) + ['000300.SH']
    print(f'\n刷新 index_factor_pro(指数技术面): {len(all_idx)} 指数(大盘+申万行业)')
    ingest_idx_factor_pro(all_idx)




# ====== refresh_industry_mapping (mapping) ======
import argparse
import os
import sys
import time
from datetime import datetime

import pandas as pd
import pymysql

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # data/→ml_training/→PKG
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'scripts'))
sys.path.insert(0, os.path.join(PKG, 'ml_training'))  # samples/data 兄弟

from tushare_token import resolve_tushare_token  # noqa: E402
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())
import tushare as ts  # noqa: E402
from samples.fetch_universe import resolve_universe  # noqa: E402

_DB = dict(host='127.0.0.1', port=3306, user='root', password='',
           database='investment_valuation', charset='utf8mb4')

# 只插分类列(metrics 留空; PB_vs_同行中位 只读 stock_code+index_code)
_INS_SQL = """INSERT INTO industry_data
(stock_code, index_code, industry_name, sw_l1_code, sw_l1_name, sw_l2_code, sw_l2_name,
 sw_l3_code, sw_l3_name, analysis_date, data_source)
VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""


def main_mapping():
    ap = argparse.ArgumentParser(description='补 industry_data 股票→行业映射(只分类, 不动系列)')
    ap.add_argument('--universe', default=None, help='placement/fullA/sample:N/file:path(与 --src 二选一)')
    ap.add_argument('--src', default=None, help='直接喂 parquet/csv/excel(股票代码 列); 优先于 --universe')
    ap.add_argument('--limit', type=int, default=0, help='只处理前 N 只缺映射股(0=全部, 冒烟用)')
    args = ap.parse_args()
    if not args.universe and not args.src:
        ap.error('需指定 --universe 或 --src')

    if args.src:
        df = pd.read_parquet(args.src) if args.src.endswith('.parquet') else pd.read_excel(args.src)
        col = '股票代码' if '股票代码' in df.columns else ('ts_code' if 'ts_code' in df.columns else 'stock_code')
        codes = [str(c).strip() for c in df[col].dropna().unique()]
    else:
        codes = resolve_universe(args.universe)['ts_code'].astype(str).tolist()
    conn = pymysql.connect(**_DB)
    mapped = set(pd.read_sql('SELECT DISTINCT stock_code FROM industry_data', conn)['stock_code'].astype(str))
    conn.close()
    todo = [c for c in codes if c not in mapped]
    if args.limit:
        todo = todo[:args.limit]
    print(f'universe={args.universe}({len(codes)}) | 已映射 {len(codes)-len(todo)} | 待补映射 {len(todo)}')
    if not todo:
        print('✅ 无待补映射股'); return

    pro = ts.pro_api()
    conn = pymysql.connect(**_DB)
    cur = conn.cursor()
    today = datetime.now().strftime('%Y%m%d')
    ok = nocls = fail = 0
    for i, c in enumerate(todo):
        try:
            df = pro.index_member_all(ts_code=c)
            if df is None or df.empty:
                nocls += 1
                continue
            # 取当前生效的(is_new=Y 优先, 否则最新 in_date)
            if 'is_new' in df.columns:
                newest = df[df['is_new'] == 'Y']
                r = newest.iloc[0] if not newest.empty else df.sort_values('in_date').iloc[-1]
            else:
                r = df.sort_values('in_date').iloc[-1]
            l3 = str(r.get('l3_code') or '').strip()
            if not l3:
                nocls += 1
                continue
            cur.execute(_INS_SQL, (
                c, l3,
                r.get('l3_name'),
                r.get('l1_code'), r.get('l1_name'),
                r.get('l2_code'), r.get('l2_name'),
                r.get('l3_code'), r.get('l3_name'),
                today, 'index_member_all'))
            ok += 1
        except Exception as e:
            fail += 1
            if fail <= 5:
                print(f'  ⚠️ {c} 失败: {e}')
        if (i + 1) % 50 == 0:
            conn.commit()
            print(f'  进度 {i+1}/{len(todo)} (补 {ok} / 无分类 {nocls} / 失败 {fail})')
        time.sleep(0.15)
    conn.commit()
    conn.close()
    print(f'完成: 补映射 {ok} / 无申万分类 {nocls} / 失败 {fail} / 共 {len(todo)}')




# ====== fetch_em_industry_stocks (em) ======

import argparse
import os
import sys
import time

# ==================== 网络环境修复 ====================

# 清除代理环境变量，避免 Privoxy/ClashX 等本地代理干扰 push2 API
for k in ['http_proxy', 'https_proxy', 'HTTP_PROXY', 'HTTPS_PROXY',
          'all_proxy', 'ALL_PROXY']:
    os.environ.pop(k, None)
os.environ['no_proxy'] = '*'
os.environ['NO_PROXY'] = '*'

# Monkey-patch akshare 的 request_with_retry，添加 trust_env=False
import random

import requests
from requests.adapters import HTTPAdapter

import akshare.utils.request as _req_mod

_original_request_with_retry = _req_mod.request_with_retry


def _patched_request_with_retry(
    url, params=None, timeout=15, max_retries=3,
    base_delay=1.0, random_delay_range=(0.5, 1.5),
):
    """带 trust_env=False 的请求，绕过系统代理。"""
    last_exception = None
    for attempt in range(max_retries):
        try:
            with requests.Session() as session:
                session.trust_env = False
                adapter = HTTPAdapter(pool_connections=1, pool_maxsize=1)
                session.mount("http://", adapter)
                session.mount("https://", adapter)
                response = session.get(url, params=params, timeout=timeout)
                response.raise_for_status()
                return response
        except requests.RequestException as e:
            last_exception = e
            if attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt) + random.uniform(*random_delay_range)
                print(f"    重试 {attempt + 1}/{max_retries}，等待 {delay:.1f}s ...")
                time.sleep(delay)
    raise last_exception


_req_mod.request_with_retry = _patched_request_with_retry

# ==================== 主逻辑 ====================

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from utils.db_manager import ValuationDB

import akshare as ak


def fetch_boards(db):
    """获取全部行业板块列表并存入数据库。"""
    print("=" * 60)
    print("步骤 1: 获取东方财富行业板块列表")
    print("=" * 60)

    for attempt in range(5):
        try:
            df = ak.stock_board_industry_name_em()
            print(f"  获取到 {len(df)} 个行业板块")
            db.save_em_industry_boards(df)
            print(f"  已保存到数据库")
            return df
        except Exception as e:
            wait = (attempt + 1) * 3
            print(f"  第 {attempt + 1} 次获取失败: {type(e).__name__}: {e}")
            if attempt < 4:
                print(f"  等待 {wait}s 后重试...")
                time.sleep(wait)

    print("  ❌ 获取行业板块列表失败，请检查网络连接")
    return None


def fetch_constituent_stocks(db, board_code, board_name):
    """获取指定行业板块的成份股并存入数据库。

    :return: 成份股数量，失败返回 -1
    """
    for attempt in range(3):
        try:
            df = ak.stock_board_industry_cons_em(symbol=board_code)
            if df is not None and len(df) > 0:
                db.save_em_industry_stocks(board_code, df)
                return len(df)
            return 0
        except Exception as e:
            if attempt < 2:
                time.sleep(2 * (attempt + 1))
    return -1


def fetch_all_stocks(db, boards_df=None):
    """获取所有行业板块的成份股。"""
    print()
    print("=" * 60)
    print("步骤 2: 获取各行业板块成份股")
    print("=" * 60)

    # 获取板块列表
    if boards_df is None:
        boards = db.get_em_industry_boards()
        if not boards:
            print("  ❌ 数据库中没有行业板块数据，请先运行 --boards-only")
            return
    else:
        boards = [
            {'board_code': row['板块代码'], 'board_name': row['板块名称']}
            for _, row in boards_df.iterrows()
        ]

    total = len(boards)
    success = 0
    failed = []
    skipped = 0

    # 查找已完成的板块（断点续采）
    existing = db.get_em_industry_boards_count()
    boards_with_stocks = set()
    if existing['boards_with_stocks'] > 0:
        conn = db.get_connection()
        try:
            cur = conn.execute(
                "SELECT board_code FROM em_industry_boards WHERE total_count > 0"
            )
            boards_with_stocks = {row[0] for row in cur.fetchall()}
        finally:
            conn.close()

    print(f"  共 {total} 个板块，已有 {len(boards_with_stocks)} 个板块的成份股数据")

    for i, board in enumerate(boards, 1):
        code = board['board_code']
        name = board['board_name']

        # 跳过已有数据的板块（除非 --force）
        if code in boards_with_stocks:
            skipped += 1
            continue

        count = fetch_constituent_stocks(db, code, name)
        if count >= 0:
            success += 1
            print(f"  [{i}/{total}] {name} ({code}): {count} 只成份股")
        else:
            failed.append(f"{name}({code})")
            print(f"  [{i}/{total}] {name} ({code}): ❌ 失败")

        # 请求间隔，避免触发限频
        time.sleep(0.5)

    print()
    print(f"  采集完成: 成功 {success}, 跳过 {skipped}, 失败 {len(failed)}")
    if failed:
        print(f"  失败板块: {', '.join(failed[:20])}")
        if len(failed) > 20:
            print(f"  ... 以及其他 {len(failed) - 20} 个")


def query_stock(db, stock_code):
    """查询股票所属行业。"""
    industries = db.get_stock_industries(stock_code)
    if industries:
        print(f"\n股票 {stock_code} 所属东方财富行业板块:")
        for ind in industries:
            print(f"  - {ind['board_name']} ({ind['board_code']})")
    else:
        print(f"\n未找到股票 {stock_code} 的行业板块数据")
        print("  请先运行采集脚本：python fetch_em_industry_stocks.py")


def show_stats(db):
    """显示数据库统计信息。"""
    stats = db.get_em_industry_boards_count()
    print("\n数据库统计:")
    print(f"  行业板块数: {stats['board_count']}")
    print(f"  已采集板块: {stats['boards_with_stocks']}")
    print(f"  成份股记录: {stats['stock_count']}")
    print(f"  不重复股票: {stats['unique_stock_count']}")

    # 展示前10个板块
    boards = db.get_em_industry_boards()
    if boards:
        print(f"\n前10个板块:")
        for b in boards[:10]:
            print(f"  {b['board_code']} {b['board_name']} ({b['total_count']} 只)")


def main_em():
    parser = argparse.ArgumentParser(description='东方财富行业板块及成份股数据采集')
    parser.add_argument('--boards-only', action='store_true',
                        help='仅采集行业板块列表，不获取成份股')
    parser.add_argument('--board', type=str,
                        help='仅采集指定板块的成份股 (如 BK1027)')
    parser.add_argument('--query', type=str,
                        help='查询股票所属行业 (如 600519)')
    parser.add_argument('--force', action='store_true',
                        help='强制重新采集（不跳过已有数据）')
    parser.add_argument('--stats', action='store_true',
                        help='显示数据库统计信息')
    args = parser.parse_args()

    db = ValuationDB()

    if args.stats:
        show_stats(db)
        return

    if args.query:
        query_stock(db, args.query)
        return

    if args.board:
        # 采集单个板块
        conn = db.get_connection()
        try:
            board = conn.execute(
                "SELECT board_code, board_name FROM em_industry_boards WHERE board_code = ?",
                (args.board,)
            ).fetchone()
        finally:
            conn.close()
        if board:
            print(f"采集板块: {board[1]} ({board[0]})")
            count = fetch_constituent_stocks(db, board[0], board[1])
            if count >= 0:
                print(f"  成功获取 {count} 只成份股")
            else:
                print(f"  ❌ 获取失败")
        else:
            print(f"未找到板块 {args.board}，请先运行全量采集")
        return

    # 全量采集
    boards_df = fetch_boards(db)

    if args.boards_only:
        if boards_df is not None:
            show_stats(db)
        return

    if boards_df is not None:
        if args.force:
            # 清空旧数据
            conn = db.get_connection()
            try:
                conn.execute("DELETE FROM em_industry_stocks")
                conn.execute("UPDATE em_industry_boards SET total_count = 0")
                conn.commit()
            finally:
                conn.close()

        fetch_all_stocks(db, boards_df)
    elif db.get_em_industry_boards_count()['board_count'] > 0:
        # 板块列表获取失败，但数据库中有旧数据，继续采集成份股
        print("\n从数据库加载已有板块列表，继续采集成份股...")
        fetch_all_stocks(db)
    else:
        print("\n❌ 无法获取行业板块数据，请检查网络后重试")

    show_stats(db)



if __name__ == '__main__':
    _cmd = sys.argv[1] if len(sys.argv) > 1 else 'refresh'
    sys.argv=[sys.argv[0]]+sys.argv[2:]
    {'refresh': main_refresh, 'mapping': main_mapping, 'em': main_em}[_cmd]()
