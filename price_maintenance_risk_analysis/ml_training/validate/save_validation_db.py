"""把回测验证结果存 DB(ml_validation + _groups + _byyear + _sections), 跨 run 可对比。

**唯一入库入口 = `save_validation_run(df, ...)`**: 逐截面 records(date,ic,long,short,ls,n)
→ 算头条 IC/ICIR/年化 + 12月组 + 按年 + 逐截面明细 → 落库。validate_5h / backtest_long_short
直接调它(不再产裸 csv)。main() 保留为 CLI 壳(读 csv → 调 save_validation_run)。

L5 报告固化层, 配合 skill quant-ml-pipeline。"""
import argparse, os, sys, warnings
warnings.filterwarnings('ignore')
import pandas as pd, numpy as np, pymysql
PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # validate/→ml_training/→PKG
for _p in (PKG, os.path.join(PKG,'ml_training'), os.path.join(PKG,'ml_training','pipeline'), os.path.join(PKG,'scripts')):
    if _p not in sys.path: sys.path.insert(0, _p)
from utils.db_manager import ValuationDB


def ann(x):  # 7m 均值 → 年化(非重叠)
    return (1 + x) ** (12 / 7) - 1


def _ensure_tables(cur):
    """幂等建 4 表 + 旧表补列。"""
    cur.execute("""CREATE TABLE IF NOT EXISTS ml_validation (
        id INT AUTO_INCREMENT PRIMARY KEY, run_date DATE, model_ver VARCHAR(128),
        horizon VARCHAR(8), sample_type VARCHAR(64), panel_path VARCHAR(256),
        period_start CHAR(8), period_end CHAR(8), random_seed INT, universe_size INT,
        method VARCHAR(256), sample_desc VARCHAR(256),
        n_rows INT, n_sections INT, n_stocks INT,
        ic_mean FLOAT, ic_std FLOAT, icir FLOAT, ic_pos_rate FLOAT,
        long_ann FLOAT, short_ann FLOAT, ls_ann_pp FLOAT,
        n_groups_pos INT, total_groups INT, mean_group_sharpe FLOAT,
        mkt_ann FLOAT, long_alpha_pp FLOAT, short_alpha_pp FLOAT,
        bull_sections INT, bull_ls FLOAT, bear_sections INT, bear_ls FLOAT,
        decision_gate VARCHAR(16), notes TEXT,
        UNIQUE KEY uq_run (model_ver, sample_type, run_date)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
    for col, ddl in [('period_start', 'CHAR(8)'), ('period_end', 'CHAR(8)'), ('random_seed', 'INT'),
                     ('universe_size', 'INT'), ('method', 'VARCHAR(256)'), ('sample_desc', 'VARCHAR(256)'),
                     ('mkt_ann', 'FLOAT'), ('long_alpha_pp', 'FLOAT'), ('short_alpha_pp', 'FLOAT'),
                     ('bull_sections', 'INT'), ('bull_ls', 'FLOAT'), ('bear_sections', 'INT'), ('bear_ls', 'FLOAT')]:
        cur.execute("""SELECT COUNT(*) FROM information_schema.COLUMNS
            WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='ml_validation' AND COLUMN_NAME=%s""", (col,))
        if cur.fetchone()[0] == 0:
            cur.execute(f"ALTER TABLE ml_validation ADD COLUMN {col} {ddl}")
    cur.execute("""CREATE TABLE IF NOT EXISTS ml_validation_groups (
        id INT AUTO_INCREMENT PRIMARY KEY, run_id INT, month_grp INT, n INT,
        long_7m FLOAT, short_7m FLOAT, ls_7m FLOAT, long_ann FLOAT, short_ann FLOAT,
        ann_diff_pp FLOAT, sharpe FLOAT, KEY idx_run (run_id)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
    cur.execute("""CREATE TABLE IF NOT EXISTS ml_validation_byyear (
        id INT AUTO_INCREMENT PRIMARY KEY, run_id INT, yr INT, n_sections INT,
        long_7m FLOAT, short_7m FLOAT, ls FLOAT, ic FLOAT, KEY idx_run (run_id)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
    cur.execute("""CREATE TABLE IF NOT EXISTS ml_validation_sections (
        id INT AUTO_INCREMENT PRIMARY KEY, run_id INT, trade_date CHAR(8), n INT,
        ic FLOAT, long_7m FLOAT, short_7m FLOAT, ls_7m FLOAT, mkt_7m FLOAT,
        UNIQUE KEY uq (run_id, trade_date), KEY idx_run (run_id)) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
    cur.execute("""SELECT COUNT(*) FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA=DATABASE() AND TABLE_NAME='ml_validation_sections' AND COLUMN_NAME='mkt_7m'""")
    if cur.fetchone()[0] == 0:
        cur.execute("ALTER TABLE ml_validation_sections ADD COLUMN mkt_7m FLOAT")


def save_validation_run(df, model_ver, sample_type, horizon=7, panel='', notes='',
                        seed=None, universe_size=None, method='', sample_desc=''):
    """逐截面 records → 落库 ml_validation*(头条 + 12月组 + 按年 + 逐截面)。唯一入库入口。
    df 需含列: date,ic,long,short,ls,n。返回 (run_id, icir, gate)。
    panel: panel parquet 路径(可选, 算 mkt_7m β基准 + 超额; 无则跳过)。"""
    df = df.copy()
    df['date'] = df['date'].astype(str)
    df['m'] = df['date'].str[4:6].astype(int)
    df['yr'] = df['date'].str[:4].astype(int)
    if panel and os.path.exists(panel):
        _p = pd.read_parquet(panel)
        _p['报价日'] = _p['报价日'].astype(str)
        _mkt = _p.groupby('报价日')['return_7m'].mean()
        df['mkt'] = df['date'].map(_mkt)
    else:
        df['mkt'] = np.nan
    ic = df['ic'].values
    lo = df['long'].values / 100; sh = df['short'].values / 100
    icir = ic.mean() / ic.std() if ic.std() > 0 else 0

    # horizon → 月数(组夏普年化用; 周→月分数; ann() 仍 7m 锚定=年化列近似, 逐截面/年原始数据准确)
    _hmap = {'1w': 0.25, '2w': 0.5, '4w': 1.0}
    try:
        h_months = _hmap.get(str(horizon), float(horizon))
    except (ValueError, TypeError):
        h_months = 7.0

    # 12 月组
    groups = []
    for m, g in df.groupby('m'):
        ls = g['ls'].values / 100; mu, sd = ls.mean(), ls.std(ddof=1)
        glo, gsh = g['long'].values / 100, g['short'].values / 100
        groups.append((int(m), len(g), round(glo.mean() * 100, 4), round(gsh.mean() * 100, 4),
                       round((glo.mean() - gsh.mean()) * 100, 4), round(ann(glo.mean()) * 100, 4),
                       round(ann(gsh.mean()) * 100, 4), round((ann(glo.mean()) - ann(gsh.mean())) * 100, 4),
                       round((mu / sd) * np.sqrt(12 / h_months), 4) if sd > 0 else 0))
    sharpes = [r[8] for r in groups]
    n_pos = sum(1 for s in sharpes if s > 0)
    gate = 'PASS' if (icir > 0.3 and n_pos >= 9 and (ann(lo.mean()) - ann(sh.mean())) > 0.05) else 'FAIL'

    # 按年
    byyear = []
    for y, g in df.groupby('yr'):
        glo, gsh = g['long'].values / 100, g['short'].values / 100
        byyear.append((int(y), len(g), round(glo.mean() * 100, 4), round(gsh.mean() * 100, 4),
                       round((glo.mean() - gsh.mean()) * 100, 4), round(g.ic.mean(), 4)))

    cfg = ValuationDB.MYSQL_CONFIG
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                           password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
    cur = conn.cursor()
    _ensure_tables(cur)

    n_stocks = int(pd.read_parquet(panel)['股票代码'].nunique()) if panel and os.path.exists(panel) else None
    if 'mkt' in df.columns and not np.all(np.isnan(df['mkt'].values)):
        _mm = float(np.nanmean(df['mkt'].values)) / 100
        mkt_ann_v = round(ann(_mm) * 100, 2)
        long_alpha = round((ann(lo.mean()) - ann(_mm)) * 100, 2)
        short_alpha = round((ann(sh.mean()) - ann(_mm)) * 100, 2)
        _bull = df[df['mkt'] > 0]; _bear = df[df['mkt'] <= 0]
        bull_n, bull_ls = len(_bull), round(float(_bull['ls'].mean()), 2) if len(_bull) else None
        bear_n, bear_ls = len(_bear), round(float(_bear['ls'].mean()), 2) if len(_bear) else None
    else:
        mkt_ann_v = long_alpha = short_alpha = bull_n = bull_ls = bear_n = bear_ls = None
    vals = (pd.Timestamp('today').date(), model_ver, horizon, sample_type, panel,
            str(df['date'].min()), str(df['date'].max()), seed, universe_size, method, sample_desc,
            len(df), df['date'].nunique(), n_stocks,
            round(float(ic.mean()), 4), round(float(ic.std()), 4), round(float(icir), 4), round(float((ic > 0).mean()), 4),
            round(ann(lo.mean()) * 100, 2), round(ann(sh.mean()) * 100, 2), round((ann(lo.mean()) - ann(sh.mean())) * 100, 2),
            n_pos, len(groups), round(float(np.mean(sharpes)), 4),
            mkt_ann_v, long_alpha, short_alpha, bull_n, bull_ls, bear_n, bear_ls,
            gate, notes)
    cur.execute("""INSERT INTO ml_validation (run_date,model_ver,horizon,sample_type,panel_path,
        period_start,period_end,random_seed,universe_size,method,sample_desc,
        n_rows,n_sections,n_stocks,
        ic_mean,ic_std,icir,ic_pos_rate,long_ann,short_ann,ls_ann_pp,n_groups_pos,total_groups,mean_group_sharpe,
        mkt_ann,long_alpha_pp,short_alpha_pp,bull_sections,bull_ls,bear_sections,bear_ls,
        decision_gate,notes)
        VALUES (%s,%s,%s,%s,%s, %s,%s,%s,%s,%s,%s, %s,%s,%s, %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,
        %s,%s,%s,%s,%s,%s,%s, %s,%s)
        ON DUPLICATE KEY UPDATE period_start=VALUES(period_start),period_end=VALUES(period_end),random_seed=VALUES(random_seed),
        universe_size=VALUES(universe_size),method=VALUES(method),sample_desc=VALUES(sample_desc),
        mkt_ann=VALUES(mkt_ann),long_alpha_pp=VALUES(long_alpha_pp),short_alpha_pp=VALUES(short_alpha_pp),
        bull_sections=VALUES(bull_sections),bull_ls=VALUES(bull_ls),bear_sections=VALUES(bear_sections),bear_ls=VALUES(bear_ls),
        icir=VALUES(icir),ls_ann_pp=VALUES(ls_ann_pp),decision_gate=VALUES(decision_gate)""", vals)
    cur.execute("SELECT id FROM ml_validation WHERE model_ver=%s AND sample_type=%s ORDER BY id DESC LIMIT 1",
                (model_ver, sample_type))
    run_id = cur.fetchone()[0]

    cur.execute("DELETE FROM ml_validation_groups WHERE run_id=%s", (run_id,))
    cur.executemany("""INSERT INTO ml_validation_groups (run_id,month_grp,n,long_7m,short_7m,ls_7m,long_ann,short_ann,ann_diff_pp,sharpe)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""", [(run_id,) + g for g in groups])
    cur.execute("DELETE FROM ml_validation_byyear WHERE run_id=%s", (run_id,))
    cur.executemany("""INSERT INTO ml_validation_byyear (run_id,yr,n_sections,long_7m,short_7m,ls,ic)
        VALUES (%s,%s,%s,%s,%s,%s,%s)""", [(run_id,) + y for y in byyear])
    cur.execute("DELETE FROM ml_validation_sections WHERE run_id=%s", (run_id,))

    def _f(x):
        return None if pd.isna(x) else float(x)
    sec = [(run_id, str(r['date']), int(r['n']), _f(r['ic']), _f(r['long']), _f(r['short']), _f(r['ls']), _f(r.get('mkt')))
           for _, r in df.iterrows()]
    cur.executemany("""INSERT INTO ml_validation_sections (run_id,trade_date,n,ic,long_7m,short_7m,ls_7m,mkt_7m)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)""", sec)
    conn.commit(); conn.close()
    print(f"✅ run_id={run_id} | {sample_type} | ICIR={icir:.3f} L-S年化={vals[20]:.2f}pp mkt年化={mkt_ann_v} {gate} | {len(groups)}组+{len(byyear)}年+{len(sec)}截面 落库", flush=True)
    return run_id, icir, gate


def register_panel(path, tag, note='', label_config=''):
    """特征 panel 以 BLOB(LONGBLOB)正经落库专用表 ml_train_wide。sha256 去重, 入库后删裸文件。
    ⚠️ **前置**: 需先 bump MySQL `max_allowed_packet` ≥ panel大小+余量(当前 64MB < 大 panel 241MB,
    会超包报错), 并重启 MySQL。配置前**不调用**(panel 暂留 ml_training/data, 落库延后)。
    配置后: `SET GLOBAL max_allowed_packet=536870912` + 重启, 再批量调本函数 backfill。返回 sha256。"""
    import hashlib
    from pyarrow import parquet as pq
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    sha = h.hexdigest()
    pf = pq.ParquetFile(path)
    n_rows, n_cols = pf.metadata.num_rows, pf.metadata.num_columns
    size_mb = os.path.getsize(path) // 1024 // 1024
    cfg = ValuationDB.MYSQL_CONFIG
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                           password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
    cur = conn.cursor()
    cur.execute("SELECT @@max_allowed_packet")
    packet = cur.fetchone()[0]
    if size_mb * 1024 * 1024 >= packet - 1_048_576:
        conn.close()
        raise RuntimeError(f"panel {size_mb}MB ≥ max_allowed_packet {packet//1024//1024}MB; 先 bump+重启 MySQL 再调用(见 docstring)")
    cur.execute("""CREATE TABLE IF NOT EXISTS ml_train_wide (  -- æ¨¡åè®­ç»å®½è¡¨(
        tag VARCHAR(128) PRIMARY KEY, kind VARCHAR(32) DEFAULT 'backtest',
        label_config VARCHAR(64), n_rows INT, n_cols INT, size_mb INT,
        sha256 CHAR(64) UNIQUE, parquet LONGBLOB, note VARCHAR(256),
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
    cur.execute("SELECT tag FROM ml_train_wide WHERE sha256=%s", (sha,))
    if cur.fetchone():
        conn.close(); print(f"⏭️ panel sha={sha[:12]} 已入库; 删裸文件", flush=True)
        os.remove(path); return sha
    with open(path, 'rb') as f:
        blob = f.read()
    cur.execute("""INSERT INTO ml_train_wide (tag,kind,label_config,n_rows,n_cols,size_mb,sha256,parquet,note)
        VALUES (%s,'backtest',%s,%s,%s,%s,%s,%s,%s)
        ON DUPLICATE KEY UPDATE n_rows=VALUES(n_rows),n_cols=VALUES(n_cols),size_mb=VALUES(size_mb),
        sha256=VALUES(sha256),parquet=VALUES(parquet),note=VALUES(note)""",
        (tag, label_config, n_rows, n_cols, size_mb, sha, blob, note))
    conn.commit(); conn.close(); os.remove(path)
    print(f"✅ panel {tag} 入库 ml_train_wide(BLOB): {n_rows}×{n_cols} {size_mb}MB sha={sha[:12]}; 裸文件已删", flush=True)
    return sha


def load_panel(tag):
    """★ DB 直读: ml_train_wide 按 tag 取 BLOB → 内存 DataFrame(无中间文件)。
    训练/LOYO/验证的样本空间唯一来源 = DB; 不再依赖本地 parquet。
    用法: df = load_panel('placement_train_20260627')  # 取代 pd.read_parquet(features_path)。"""
    import io
    cfg = ValuationDB.MYSQL_CONFIG
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                           password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
    cur = conn.cursor()
    cur.execute("SELECT parquet FROM ml_train_wide WHERE tag=%s", (tag,))
    row = cur.fetchone(); conn.close()
    if not row:
        raise KeyError(f"ml_train_wide 无 tag={tag}")
    import pandas as pd
    return pd.read_parquet(io.BytesIO(row[0]))


def load_features(src):
    """统一装载: src 是 ml_train_wide tag → load_panel(DB直读); 否则当文件路径 → read_parquet。
    训练脚本可用此替代 pd.read_parquet(features_path), 既支持 DB tag 也兼容旧文件路径。"""
    import pandas as pd
    # 探: src 是否为 ml_train_wide 已知 tag
    cfg = ValuationDB.MYSQL_CONFIG
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                           password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM ml_train_wide WHERE tag=%s", (src,))
    is_tag = cur.fetchone() is not None; conn.close()
    if is_tag:
        return load_panel(src)
    return pd.read_parquet(src)


def restore_panel(tag, out_path=None):
    """从 ml_train_wide(模型训练宽表)按 tag 还原 panel parquet 到 out_path。
    DB 为样本空间唯一来源后, 训练/LOYO/验证经此取回 panel 到本地再读。
    不传 out_path 则写到 PKG/ml_training/data/<tag>.parquet。返回写出路径。"""
    cfg = ValuationDB.MYSQL_CONFIG
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                           password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
    cur = conn.cursor()
    cur.execute("SELECT parquet, n_rows, n_cols FROM ml_train_wide WHERE tag=%s", (tag,))
    row = cur.fetchone()
    conn.close()
    if not row:
        raise KeyError(f"ml_train_wide 无 tag={tag}")
    blob, n_rows, n_cols = row
    if out_path is None:
        out_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                                'ml_training', 'data', f'{tag}.parquet')
    with open(out_path, 'wb') as f:
        f.write(blob)
    print(f"✅ 还原 panel {tag} → {out_path} ({n_rows}×{n_cols}, {len(blob)//1024//1024}MB)", flush=True)
    return out_path


def list_panels():
    """列出 ml_train_wide 所有样本空间(tag/行/列/大小/sha)。"""
    cfg = ValuationDB.MYSQL_CONFIG
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                           password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
    cur = conn.cursor()
    cur.execute("SELECT tag, n_rows, n_cols, size_mb, LEFT(sha256,10), label_config FROM ml_train_wide ORDER BY tag")
    rows = cur.fetchall(); conn.close()
    print(f"ml_train_wide(模型训练宽表) {len(rows)} 个样本空间:")
    for tag, nr, nc, sz, sha, lc in rows:
        print(f"  {tag} | {nr}×{nc} {sz}MB sha={sha} | {lc}")
    return rows


def main():
    ap = argparse.ArgumentParser(description='回测验证结果落库(读 csv → save_validation_run)')
    ap.add_argument('--csv', required=True, help='逐截面 records(date,ic,long,short,ls,n)')
    ap.add_argument('--model-ver', required=True)
    ap.add_argument('--sample-type', required=True, help='pilot_500_stratified / full_A / placement / backtest_1500')
    ap.add_argument('--horizon', default='7', help='期限(7/3/1/2w/1w, 默认7)')
    ap.add_argument('--panel', default='')
    ap.add_argument('--notes', default='')
    ap.add_argument('--seed', type=int, default=None)
    ap.add_argument('--universe-size', type=int, default=None)
    ap.add_argument('--method', default='')
    ap.add_argument('--sample-desc', default='')
    args = ap.parse_args()
    df = pd.read_csv(args.csv)
    save_validation_run(df, args.model_ver, args.sample_type, args.horizon, args.panel, args.notes,
                        args.seed, args.universe_size, args.method, args.sample_desc)


if __name__ == '__main__':
    main()
