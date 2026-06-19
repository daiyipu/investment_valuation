#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""因子效力评估(coverage-aware + bar粒度↔target期限对齐)。

核心理念: 因子的 K 线粒度要匹配 target 期限, 否则错配假阴性。
对齐规则(可在 ALIGN 调):
  日线(无后缀) → 短标签 标签_短_{1w/2w/4w}_gray(p25/p75 极性)
  周线 _W     → 中期 标签_极性_灰度剔除_{1m/3m}m(-20/+10)
  月线 _M     → 长期 标签_极性_灰度剔除_{3m/7m}m(-20/+10)
coverage-aware: 每个因子只在自身非空样本上评(不 median 填 pooled, 避假阴性)。

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
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # PKG(utils)
from train_scorecard import calc_iv_all_features   # noqa: E402
from sklearn.metrics import roc_auc_score          # noqa: E402

# bar 粒度 → 对齐的 target 期限。'1w'/'2w'/'4w'=短线gray; '1m'/'3m'/'7m'=gray(-20/+10)。
ALIGN = {
    'daily': ['1w', '2w', '4w', '1m'],   # 日线: 短标签 + 1m
    '_W':    ['1m', '3m'],               # 周线: 中期
    '_M':    ['3m', '7m'],               # 月线: 长期
}
RETURN_COLS = ['return_1m', 'return_3m', 'return_7m', 'return_1w', 'return_2w', 'return_4w']
SHORT = {'1w': 'return_1w', '2w': 'return_2w', '4w': 'return_4w'}
MONTH = {'1m': 'return_1m', '3m': 'return_3m', '7m': 'return_7m'}


def granularity(col):
    if col.endswith('_W'):
        return '_W'
    if col.endswith('_M'):
        return '_M'
    return 'daily'


def build_labels(df):
    """内存构 gray 标签列: 标签_短_{w}_gray(p25/p75), 标签_极性_灰度剔除_{h}m(-20/+10)。"""
    for w, rc in SHORT.items():
        s = pd.to_numeric(df[rc], errors='coerce')
        p25, p75 = s.quantile(0.25), s.quantile(0.75)
        df[f'_lab_{w}'] = np.where(s > p75, 1, np.where(s < p25, 0, np.nan))
    for h, rc in MONTH.items():
        r = pd.to_numeric(df[rc], errors='coerce')
        df[f'_lab_{h}'] = np.where(r > 10, 1, np.where(r < -20, 0, np.nan))
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
    ap.add_argument('--horizons', default='aligned',
                    help="aligned(默认,按粒度对齐) | all(全期限1w/2w/4w/1m/3m/7m) | 自定义如 1m,3m")
    args = ap.parse_args()

    # 确定因子列
    if args.prefix:
        import pymysql
        from utils.db_manager import ValuationDB
        conn = pymysql.connect(**ValuationDB.MYSQL_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                    "WHERE TABLE_SCHEMA=%s AND TABLE_NAME='placement_evaluation'", (ValuationDB.MYSQL_CONFIG['database'],))
        factors = [r[0] for r in cur.fetchall() if r[0].startswith(args.prefix)]
        conn.close()
    else:
        factors = args.factors
    if not factors:
        print('❌ 未指定因子(用 位置参数 或 --prefix)'); sys.exit(1)

    df = build_labels(load(factors))
    all_h = ['1w', '2w', '4w', '1m', '3m', '7m']
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
