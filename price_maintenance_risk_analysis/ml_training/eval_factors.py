#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""因子效力评估(coverage-aware + bar粒度↔target期限对齐)。

核心理念: 因子的 K 线粒度要匹配 target 期限, 否则错配假阴性。
对齐规则(可在 ALIGN 调): 全部用生产口径 GRAY_CFG(gray, 各期限 sweep 定阈值)。
  日线(无后缀) → 1w/2w/1m   周线 _W → 1m/3m   月线 _M → 3m/7m
coverage-aware: 每个因子只在自身非空样本上评(不 median 填 pooled, 避假阴性)。
⚠ 2026-06-22: quantile p25/p75 短标签方案已移除(与生产 gray 口径不一致、难操作);
   1w/2w 改用生产 GRAY_CFG; 4w 无生产模型, 去除。

数据源: placement_evaluation DB(因子列 + return_*m + return_*w); gray 标签内存算。
复用: train_scorecard.calc_iv_all_features(IV) + sklearn roc_auc_score(AUC)。

用法:
  python ml_training/eval_factors.py smc_ote smc_liqvoid       # 指定因子(对齐评估)
  python ml_training/eval_factors.py --prefix smc_             # 所有 smc_* 因子
  python ml_training/eval_factors.py chip_concentration --horizons all  # 全期限(不对齐)
"""
import argparse
import sys
import os

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))            # ml_training/
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pipeline'))  # 管线模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # PKG(utils)
from train_scorecard import calc_iv_all_features   # noqa: E402
from sklearn.metrics import roc_auc_score          # noqa: E402

# bar 粒度 → 对齐的 target 期限。全部生产口径 GRAY_CFG(gray); quantile p25/p75 已移除。
ALIGN = {
    'daily': ['1w', '2w', '1m'],   # 日线: 周 + 1m
    '_W':    ['1m', '3m'],         # 周线: 中期
    '_M':    ['3m', '7m'],         # 月线: 长期
}
RETURN_COLS = ['return_1m', 'return_3m', 'return_7m', 'return_1w', 'return_2w']
# horizon → (DB return 列, GRAY_CFG 键); 全部生产 gray 口径
HORIZON_RET = {'1w': 'return_1w', '2w': 'return_2w', '1m': 'return_1m', '3m': 'return_3m', '7m': 'return_7m'}
HORIZON_GK = {'1w': '1w', '2w': '2w', '1m': 1, '3m': 3, '7m': 7}


def granularity(col):
    if col.endswith('_W'):
        return '_W'
    if col.endswith('_M'):
        return '_M'
    return 'daily'


def build_labels(df):
    """内存构 gray 标签列(全部生产口径 GRAY_CFG, 与训练一致)。

    1w/2w/1m/3m/7m 各用 GRAY_CFG sweep 定的阈值。quantile p25/p75 方案已移除(2026-06-22)。
    """
    from train_horizon_models import GRAY_CFG
    for h, rc in HORIZON_RET.items():
        lo, hi = GRAY_CFG[HORIZON_GK[h]]
        r = pd.to_numeric(df[rc], errors='coerce')
        df[f'_lab_{h}'] = np.where(r > hi, 1, np.where(r < lo, 0, np.nan))
    return df


def load(factor_cols):
    import pymysql
    from utils.db_manager import ValuationDB
    conn = pymysql.connect(**ValuationDB.MYSQL_CONFIG, cursorclass=pymysql.cursors.DictCursor)
    cur = conn.cursor()
    cols = ', '.join(f'`{c}`' for c in factor_cols + RETURN_COLS)
    cur.execute(f'SELECT stock_code, issue_date, {cols} FROM placement_evaluation '
                'WHERE issue_date IS NOT NULL AND LENGTH(issue_date)=8')
    df = pd.DataFrame(cur.fetchall())
    conn.close()
    return df


def load_parquet(path):
    """从 features_derived.parquet 读 derive 因子, 全部 horizon 用生产口径 GRAY_CFG 从收益重算 gray。

    不用 parquet 预算列(预算列口径可能与生产不一致); 1w/2w/1m/3m/7m 统一从 GRAY_CFG 重算。
    """
    from train_horizon_models import GRAY_CFG
    df = pd.read_parquet(path)
    parq_rc = {'1w': '1周涨跌幅', '2w': '2周涨跌幅', '1m': '1个月涨跌幅', '3m': '3个月涨跌幅', '7m': '7个月涨跌幅'}
    for h, rc in parq_rc.items():
        lo, hi = GRAY_CFG[HORIZON_GK[h]]
        r = pd.to_numeric(df[rc], errors='coerce') if rc in df.columns else pd.Series(np.nan, index=df.index)
        df[f'_lab_{h}'] = np.where(r > hi, 1, np.where(r < lo, 0, np.nan))
    return df


def eval_factor(df, col, horizons):
    """coverage-aware IV/AUC: col × 各 horizon gray 标签, 只在 col 非空样本上。"""
    rows = []
    for h in horizons:
        lab = f'_lab_{h}'
        sub = df[[col, lab]].dropna()
        sub = sub[pd.to_numeric(sub[col], errors='coerce').notna()]
        if len(sub) < 200:
            continue
        yy = sub[lab].astype(int)
        x = pd.to_numeric(sub[col], errors='coerce')
        iv = calc_iv_all_features(pd.DataFrame({col: x}), yy).iloc[0]['iv']
        try:
            auc = roc_auc_score(yy, x)
        except Exception:
            auc = float('nan')
        rows.append({'horizon': h, 'n': len(sub), 'IV': iv, 'AUC': auc})
    return rows


def main():
    ap = argparse.ArgumentParser(description='因子效力评估(coverage-aware + 粒度对齐)')
    ap.add_argument('factors', nargs='*', help='因子列名(可多个)')
    ap.add_argument('--prefix', help='评估所有以该前缀开头的因子列(如 smc_)')
    ap.add_argument('--parquet', help='从 features_derived.parquet 读 derive 因子+预计算灰度标签(评 量价/turnover 等)')
    ap.add_argument('--horizons', default='aligned',
                    help="aligned(默认,按粒度对齐) | all(全期限1w/2w/1m/3m/7m) | 自定义如 1m,3m")
    args = ap.parse_args()

    # 确定因子列 + 载入
    if args.parquet:
        df = load_parquet(args.parquet)
        if args.prefix:
            factors = [c for c in df.columns if isinstance(c, str) and c.startswith(args.prefix)]
        else:
            factors = args.factors
        miss = [f for f in factors if f not in df.columns]
        if miss:
            print(f'⚠️ parquet 缺列(跳过): {miss}')
            factors = [f for f in factors if f in df.columns]
    elif args.prefix:
        import pymysql
        from utils.db_manager import ValuationDB
        conn = pymysql.connect(**ValuationDB.MYSQL_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                    "WHERE TABLE_SCHEMA=%s AND TABLE_NAME='placement_evaluation'", (ValuationDB.MYSQL_CONFIG['database'],))
        factors = [r[0] for r in cur.fetchall() if r[0].startswith(args.prefix)]
        conn.close()
        df = build_labels(load(factors))
    else:
        factors = args.factors
        df = build_labels(load(factors))
    if not factors:
        print('❌ 未指定因子(用 位置参数 或 --prefix)'); sys.exit(1)

    all_h = ['1w', '2w', '1m', '3m', '7m']
    print(f'因子 {len(factors)} 个 | 样本 {len(df)} | 对齐: {args.horizons}\n')
    for col in factors:
        g = granularity(col)
        if args.horizons == 'aligned':
            hs = ALIGN[g]
        elif args.horizons == 'all':
            hs = all_h
        else:
            hs = [h.strip() for h in args.horizons.split(',')]
        rows = eval_factor(df, col, hs)
        if not rows:
            print(f'[{col}] ({g}) 样本不足, 跳过\n'); continue
        print(f'[{col}]  bar={g}  对齐期限={hs}')
        for r in rows:
            tag = ' (反)' if r['AUC'] < 0.48 else ''
            flag = '⭐' if r['IV'] >= 0.05 else ('·' if r['IV'] >= 0.01 else '✗')
            print(f"  {flag} {r['horizon']:4s} n={r['n']:5d}  IV={r['IV']:.4f}  AUC={r['AUC']:.3f}{tag}")
        print()


if __name__ == '__main__':
    main()
