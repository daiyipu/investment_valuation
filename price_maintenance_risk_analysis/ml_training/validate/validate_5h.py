"""一次性: 给 panel 补 return_3m/1m/1w/2w 标签 + 5期限 IC/ICIR/L-S 验证。
标签=前瞻收益(非特征), post-process panel, 不动 build_features 入口(三层一入口不违)。
月标签复用 backtest_long_short.fwd_returns(months); 周标签=交易日数前移(_fwd_ndays)。
用法: python /tmp/validate_5h.py <panel.parquet>
"""
import os, sys, pickle, argparse
from datetime import datetime
PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(PKG)
sys.path.insert(0, PKG)
sys.path.insert(0, os.path.join(PKG, 'scripts'))
import numpy as np, pandas as pd
from scipy.stats import spearmanr
sys.path.insert(0, os.path.join(PKG, 'ml_training'))
sys.path.insert(0, os.path.join(PKG, 'ml_training', 'pipeline'))
from validate import backtest_long_short as bls           # fwd_returns + _CLOSE_CACHE + score_sc 路径
from deploy.db_model_store import load_predict_bundle
from deploy.predict_profitability import score_sc  # SC 打分(同 backtest_long_short)
from validate.save_validation_db import save_validation_run  # 验证结果直接入库(不产裸csv); panel落库延后

# 各期限最新 SC 模型(确定性指定, 非 latest)
MODELS = {
    '7':  'v_sc_20260624_1452_7m_gray_sc_15feat',
    '3':  'v_sc_20260623_1723_3m_gray_sc_11feat',
    '1':  'v_sc_20260622_0158_1m_gray_sc_18feat',
    '2w': 'v_sc_20260622_1141_2w_gray_sc_13feat',
    '1w': 'v_sc_20260622_1125_1w_gray_sc_21feat',
}
WEEK_TD = {'1w': 5, '2w': 10}   # 1w=5 交易日, 2w=10 交易日


def _fwd_ndays(series, date_yyyymmdd, n_days):
    """周标签前移收益%。复用 compute_labels.ingest_shortterm 同款口径:
    - 窗口 = WINDOWS{'1w':5,'2w':10}(交易日), 不自发明。
    - i0 = 报价日当天/前一交易日(bisect, 同 ingest_shortterm)。
    - 退市/停牌不足 n_days → 用最后可得价计亏(同 backtest_long_short._bench_with_delist
      的 PIT 诚实策略, 避免生存偏差; 与 panel 月标签 return_7m 口径一致)。
    series = build_series 产出的 [(ordinal,close)...] 升序。"""
    if not series:
        return None
    try:
        t0 = datetime.strptime(date_yyyymmdd, '%Y%m%d').toordinal()
    except ValueError:
        return None
    # 找 t0 当日/最近前一个交易日的价(入场价)
    import bisect
    dates = [s[0] for s in series]
    i0 = bisect.bisect_right(dates, t0) - 1   # ≤t0 的最后一个
    if i0 < 0:
        return None
    c0 = series[i0][1]
    if c0 == 0:
        return None
    # 第 n_days 个交易日之后(i0 + n_days, 不超过末尾)
    i1 = min(i0 + n_days, len(series) - 1)
    c1 = series[i1][1]
    if series[i1][0] <= t0:    # t0 之后无交易
        return None
    return (c1 / c0 - 1) * 100


def _warm_close_cache(stocks):
    """从 DB stock_qfq_daily 批量读 close → 预热 bls._CLOSE_CACHE(免逐股 pro_bar 限频)。
    同 build_backtest_panel 的 _qfq_bulk_read 取数方式, chunk 800/批。"""
    import pymysql
    from data.compute_labels import build_series
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    stocks = sorted(set(stocks))
    n = 0
    for i in range(0, len(stocks), 800):
        blk = stocks[i:i + 800]
        ph = ','.join(['%s'] * len(blk))
        sql = f"SELECT stock_code,trade_date,close FROM stock_qfq_daily WHERE stock_code IN ({ph})"
        chunk = pd.read_sql(sql, conn, params=blk)
        for c, g in chunk.groupby('stock_code'):
            g = g.sort_values('trade_date')
            cm = dict(zip(g['trade_date'].astype(str), pd.to_numeric(g['close'], errors='coerce')))
            bls._CLOSE_CACHE[c] = build_series(cm)
            n += 1
        print(f'  预热 close {min(i+800,len(stocks))}/{len(stocks)} (累计{n}股)', flush=True)
    conn.close()
    return n


def add_labels(panel):
    """补 return_3m/1m/1w/2w(7m 已由 build 入)。DB 批量预热 close → 月/周标签都走热缓存。"""
    panel['报价日'] = panel['报价日'].astype(str)
    scode_col = '股票代码' if '股票代码' in panel.columns else 'stock_code'
    print('  预热 close 缓存(DB 批量读 stock_qfq_daily)...', flush=True)
    _warm_close_cache(panel[scode_col].unique().tolist())
    for col in ['return_3m', 'return_1m', 'return_1w', 'return_2w']:
        if col not in panel.columns:
            panel[col] = np.nan
    sec = panel.groupby('报价日')
    nsec = len(sec)
    for i, (d, g) in enumerate(sec, 1):
        codes = g[scode_col].tolist()
        for h, col in [(3, 'return_3m'), (1, 'return_1m')]:
            fwd = bls.fwd_returns(codes, d, months=h)   # _CLOSE_CACHE 已热, 免 pro_bar
            panel.loc[g.index, col] = [fwd.get(c) for c in codes]
        for tag, nd in WEEK_TD.items():
            col = f'return_{tag}'
            vals = [_fwd_ndays(bls._CLOSE_CACHE.get(c), d, nd) for c in codes]
            panel.loc[g.index, col] = vals
        if i % 10 == 0 or i == nsec:
            print(f'  标签 {i}/{nsec} 截面', flush=True)
    return panel


def validate(panel, horizon, model_ver, ret_col):
    """单期限: 逐截面 score_sc → IC/L-S。返回 records DataFrame + 摘要。"""
    bundle = pickle.loads(load_predict_bundle(model_ver)['lr_bundle'])
    feats = bundle['features']
    miss = [f for f in feats if f not in panel.columns]
    if miss:
        print(f'  ⚠️ {horizon}: panel 缺 {len(miss)} 特征 {miss[:5]}...')
    recs = []
    for d, g in panel.groupby('报价日'):
        if len(g) < 50:
            continue
        s = g[ret_col]
        valid = s.notna()
        if valid.sum() < 50:
            continue
        proba, _ = score_sc(bundle, g[feats].copy())
        p = pd.Series(proba, index=g.index)[valid]
        s = s[valid]
        common = p.index.intersection(s.index)
        p, s = p.loc[common], s.loc[common]
        if len(s) < 50:
            continue
        ic = float(spearmanr(p, s).correlation)
        k = max(1, int(len(s) * 0.10))
        order = p.sort_values()
        short = float(s.loc[order.index[:k]].mean())
        long = float(s.loc[order.index[-k:]].mean())
        recs.append({'date': d, 'ic': ic, 'long': long, 'short': short,
                     'ls': long - short, 'n': int(len(s))})
    if not recs:
        print(f'  ❌ {horizon}: 无有效截面'); return None
    df = pd.DataFrame(recs)
    ic = df['ic'].values
    icir = ic.mean() / ic.std() if ic.std() > 0 else 0
    ls = df['ls'].values
    pos = int((ls > 0).sum())
    print(f'  {horizon:>3} ({model_ver[-22:]}): n截面={len(df)} IC均值={ic.mean():+.4f} '
          f'ICIR={icir:+.3f} | L-S均值={ls.mean():+.2f}% ({pos}/{len(df)}正) | '
          f'多头={df["long"].mean():+.2f}% 最差={df["short"].mean():+.2f}%')
    return df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('panel')
    ap.add_argument('--tag', default='1500')
    args = ap.parse_args()
    print(f'读 panel: {args.panel}', flush=True)
    panel = pd.read_parquet(args.panel)
    print(f'  {len(panel)} 行 × {panel.shape[1]} 列, {panel["报价日"].nunique()} 截面', flush=True)

    panel = add_labels(panel)
    # 覆盖率
    for col in ['return_7m', 'return_3m', 'return_1m', 'return_2w', 'return_1w']:
        if col in panel.columns:
            print(f'  {col} 覆盖 {panel[col].notna().mean() * 100:.1f}%')
    out_panel = args.panel.replace('.parquet', '_5labels.parquet')
    panel.to_parquet(out_panel, index=False)
    print(f'  带5标签 panel 写出: {out_panel} (panel BLOB 落库延后至 MySQL max_allowed_packet 配好)', flush=True)

    print('\n' + '=' * 70)
    print('5 期限验证(IC/ICIR/多空) → 直接入库 ml_validation*(不产裸 csv)')
    print('=' * 70)
    ret_cols = {'7': 'return_7m', '3': 'return_3m', '1': 'return_1m',
                '2w': 'return_2w', '1w': 'return_1w'}
    summary = {}
    for h in ['7', '3', '1', '2w', '1w']:
        df = validate(panel, h, MODELS[h], ret_cols[h])
        if df is not None:
            summary[h] = {'icir': df.ic.mean() / df.ic.std() if df.ic.std() > 0 else 0,
                          'ic': df.ic.mean(), 'ls': df.ls.mean()}
            # 7m 带 panel 算 mkt 基准; 其余期限不传 panel 避免 7m 锚定错 β
            save_validation_run(df, MODELS[h], f'backtest_{args.tag}', h,
                                panel=out_panel if h == '7' else '',
                                universe_size=5603, method='全A随机股月末截面SC打分IC/L-S(恢复414特征)')
    # panel BLOB 落库(register_panel)延后: 需先 bump max_allowed_packet≥512MB + 重启 MySQL, 配好后再 backfill
    print('=' * 70)
    print('汇总:')
    for h, s in summary.items():
        print(f'  {h:>3}: IC={s["ic"]:+.4f} ICIR={s["icir"]:+.3f} L-S={s["ls"]:+.2f}%')


if __name__ == '__main__':
    main()
