#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""评分卡概率拆解 — 解释单只股票为何得分高/低(只读, 不改模型)。

机制回顾:
  评分卡 = WOE 分箱 + 逻辑回归。
  每个特征按训练集分位切成 ~5 箱, 每箱 WOE = ln(好样本占比 / 坏样本占比)。
  正 WOE → 该箱历史上"赢家"多; 负 WOE → "输家"多。
  LR 在 WOE 编码后的特征上学习, 概率 = sigmoid( 截距 + Σ(系数ᵢ × 该箱WOEᵢ) )。
  故「扣分项」= 系数 × WOE 最负的几项;「加分项」= 最正的几项。

用法:
  python ml_training/explain_scorecard.py <股票代码> [股票简称] [报价日YYYYMMDD]
  python ml_training/explain_scorecard.py 300604.SZ 长川科技 20240601
  python ml_training/explain_scorecard.py 300604.SZ              # 报价日=今天
  python ml_training/explain_scorecard.py 300604.SZ 20240601 --blue   # 拆 BLUE(4基本面)模型

说明: GREEN=当前生产(11技术特征, current.full); BLUE=4基本面回退模型。
      先跑批量全流程(单行临时Excel)重建特征, 与 predict 特征空间完全一致。
"""

import os
import sys
import glob
import pickle
import subprocess
import tempfile
import warnings
import unicodedata

import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # ml_training/(搬自 report/ 升一层)           # ml_training/
PKG_DIR = os.path.dirname(SCRIPT_DIR)                             # price_maintenance_risk_analysis/
SCRIPTS_DIR = os.path.join(PKG_DIR, 'scripts')
BATCH_SCRIPT = os.path.join(SCRIPTS_DIR, 'batch_screen_and_score.py')
sys.path.insert(0, SCRIPT_DIR)
sys.path.insert(0, os.path.join(SCRIPT_DIR, 'pipeline'))   # 管线模块已移入 pipeline/

VNPY_PYTHON = os.path.expanduser('~/anaconda3/envs/vnpy/bin/python')
PYTHON = VNPY_PYTHON if os.path.exists(VNPY_PYTHON) else sys.executable


# ── 终端对齐 ──
def _cw(s):
    return sum(2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1 for ch in str(s))


def _fv(v):
    """数值格式化: NaN→'缺失(中位数填)', 否则 3 位有效。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return '缺失'
    if f != f:   # NaN
        return '缺失'
    if f == 0:
        return '0'
    if abs(f) >= 1000 or abs(f) < 0.001:
        return f'{f:.2e}'
    return f'{f:.4g}'


def build_scored_excel(code, name, issue_date):
    """跑批量全流程(单行临时输入)→ 返回 scored Excel 路径(在临时目录内)。"""
    with tempfile.TemporaryDirectory(prefix='explain_sc_') as td:
        inp = os.path.join(td, 'input.xlsx')
        pd.DataFrame([{'股票代码': code, '股票简称': name, '报价日': issue_date}]) \
            .to_excel(inp, index=False)
        cmd = [PYTHON, BATCH_SCRIPT, '--input', inp, '--sheet', '0', '--force']
        print('⏳ 重建特征中(跑批量全流程, 约 30~90 秒)...\n')
        proc = subprocess.run(cmd, cwd=PKG_DIR, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        if proc.returncode != 0:
            print('❌ 批量流程失败'); sys.exit(1)
        hits = glob.glob(os.path.join(td, '*_scored.xlsx'))
        if not hits:
            print('❌ 未生成 scored Excel'); sys.exit(1)
        scored = hits[0]
        # 复制到进程外(临时目录退出即删)
        keep = tempfile.mktemp(suffix='_scored.xlsx')
        import shutil
        shutil.copy(scored, keep)
        return keep


def build_feature_df(scored_path):
    """复刻 predict_profitability ML-1~3: 读 Excel + DB + 衍生 + 因子引擎 → 单行全特征 df。"""
    from features.export_features import load_db_features, load_scored_features, load_financial_ratios
    from features.derive_features import (
        derive_fcf_growth_rates, derive_fcf_cross_metrics,
        derive_financial_score_deltas, derive_valuation_relative,
        derive_market_momentum, derive_industry_valuation_growth,
        derive_market_index_features, derive_strategy_signals, derive_alpha_beta_factors,
    )

    scored = load_scored_features(scored_path)
    sample_keys = []
    for _, row in scored.iterrows():
        code = row['股票代码']
        d = str(row.get('报价日', '')).replace('.0', '').strip()
        d = d if d and d != 'nan' and len(d) >= 8 else None
        sample_keys.append((code, d))

    db_feats = load_db_features(sample_keys)
    try:
        ratio_feats = load_financial_ratios(sample_keys)
    except Exception:
        ratio_feats = pd.DataFrame()

    scored = scored.reset_index(drop=True)
    _dup = (set(scored.columns) & set(db_feats.columns)) - {'股票代码'}
    db_cols = [c for c in db_feats.columns if c != '股票代码' and c not in _dup]
    df = pd.concat([scored, db_feats[db_cols].reset_index(drop=True)], axis=1)
    if not ratio_feats.empty:
        df = pd.concat([df.reset_index(drop=True), ratio_feats.reset_index(drop=True)], axis=1)

    str_keep = {'股票代码', '股票简称', '最终结论', '一级行业', '二级行业', '三级行业',
                '定价方式', '定增决策', '行业代码', '行业名称'}
    for c in df.columns:
        if c in str_keep:
            df[c] = df[c].astype(str)
        elif df[c].dtype == object or str(df[c].dtype) == 'category':
            df[c] = pd.to_numeric(df[c], errors='coerce')

    for func in [derive_fcf_growth_rates, derive_fcf_cross_metrics,
                 derive_financial_score_deltas, derive_valuation_relative, derive_market_momentum]:
        df = func(df)
    try:
        df = derive_industry_valuation_growth(df)
        df = derive_market_index_features(df)
    except Exception:
        pass
    try:
        df = derive_strategy_signals(df)
        df = derive_alpha_beta_factors(df)
    except Exception:
        pass
    df = df.replace([np.inf, -np.inf], np.nan)
    return df


def explain_model(df, version, tag):
    """对一个 SC 模型版本, 拆解每特征 logit 贡献并打印。"""
    from deploy.db_model_store import load_predict_bundle
    from eval_loyo import apply_woe

    bundle = load_predict_bundle(version)
    sc = pickle.loads(bundle['lr_bundle'])
    feats = sc['features']
    medians = sc.get('medians', {})
    woe_bins = sc['woe_bins']
    lr = sc['lr_model']
    coefs = dict(zip(feats, lr.coef_[0]))
    intercept = float(lr.intercept_[0])

    # 构建单行特征矩阵(缺失按训练 median 填, 与 predict 一致)
    row0 = df.iloc[0]
    X = pd.DataFrame(index=[0])
    for f in feats:
        X[f] = pd.to_numeric(row0.get(f), errors='coerce') if f in df.columns else np.nan
    for f in feats:
        X[f] = X[f].fillna(medians.get(f, 0))
    X = X.replace([np.inf, -np.inf], 0)
    Xw = apply_woe(X, feats, woe_bins).replace([np.inf, -np.inf], 0).fillna(0)

    # 概率自洽校验
    logit_parts = []
    for f in feats:
        logit_parts.append(coefs[f] * float(Xw[f].iloc[0]))
    logit = intercept + sum(logit_parts)
    proba = 1 / (1 + np.exp(-logit))

    print('\n' + '═' * 70)
    print(f'  {tag}  版本 {version}  ({len(feats)} 特征)')
    print(f'  logit = {intercept:+.3f}(截距) + Σ贡献 = {logit:+.3f}  →  P = {proba*100:.1f}%')
    print('═' * 70)

    # 组装并按贡献排序
    records = []
    for f in feats:
        raw = row0.get(f, np.nan)
        try:
            raw_f = float(raw)
            raw_na = raw_f != raw_f
        except (TypeError, ValueError):
            raw_f = np.nan; raw_na = True
        woe = float(Xw[f].iloc[0])
        coef = coefs[f]
        contrib = coef * woe
        # 原始值是否被 median 填充过
        filled = raw_na
        records.append({'feat': f, 'raw': raw_f, 'filled': filled,
                        'woe': woe, 'coef': coef, 'contrib': contrib})

    dfc = pd.DataFrame(records).sort_values('contrib')

    def _show(dfc_slice, title):
        print(f'\n  【{title}】')
        hdr = f"  {'特征':<22}{'原始值':>12}  {'箱WOE':>8}  {'系数':>8}  {'logit贡献':>11}"
        print(hdr)
        print('  ' + '─' * 66)
        for _, r in dfc_slice.iterrows():
            raw = ('缺失→填中位数' if r['filled'] else _fv(r['raw']))
            print(f"  {r['feat']:<22}{raw:>16}  {r['woe']:+8.3f}  {r['coef']:+8.3f}  {r['contrib']:+11.3f}")

    neg = dfc[dfc['contrib'] < 0]
    pos = dfc[dfc['contrib'] > 0]
    _show(neg.head(8), '扣分项(负贡献, 推低概率)')
    _show(pos.sort_values('contrib', ascending=False).head(8), '加分项(正贡献, 推高概率)')

    # 解读: 哪些是"落到历史输家箱"
    print('\n  【解读】')
    bad_bin = dfc[(dfc['woe'] < -0.3) & (dfc['coef'] > 0)]
    if len(bad_bin):
        print('  落入「历史输家箱」(WOE<-0.3) 且模型正向加权(系数>0)的特征 → 主因:')
        for _, r in bad_bin.head(5).iterrows():
            raw = ('缺失' if r['filled'] else _fv(r['raw']))
            print(f"    • {r['feat']}: 原始值 {raw}, 箱WOE={r['woe']:+.2f}, 贡献 {r['contrib']:+.2f}")


def main():
    import argparse
    ap = argparse.ArgumentParser(
        description='评分卡概率拆解(只读解释, 不改模型)',
        usage='%(prog)s <股票代码> [股票简称] [报价日YYYYMMDD] [--blue]',
    )
    ap.add_argument('code')
    ap.add_argument('name', nargs='?', default='')
    ap.add_argument('issue_date', nargs='?', default=None)
    ap.add_argument('--blue', action='store_true', help='只拆 BLUE(4基本面)模型; 默认拆 GREEN(当前生产)')
    args = ap.parse_args()

    code = args.code.strip()
    issue_date = (args.issue_date or __import__('datetime').datetime.now().strftime('%Y%m%d')).strip()

    scored = build_scored_excel(code, args.name, issue_date)
    try:
        df = build_feature_df(scored)
        if df.empty:
            print('❌ 特征构建为空'); sys.exit(1)

        from deploy.model_registry import get_current
        from deploy.db_model_store import list_model_metas

        green_ver = get_current('full')
        metas = list_model_metas(kind='gray')
        cur_nfeat = 0
        try:
            from deploy.db_model_store import load_predict_bundle
            cur_nfeat = len(load_predict_bundle(green_ver)['features'])
        except Exception:
            pass
        blue_cands = [m for m in metas if m['version'] != green_ver and not m.get('lgb_model')
                      and m.get('kind') == 'gray']
        blue_cands.sort(key=lambda m: abs(m.get('n_features', 0) - cur_nfeat), reverse=True)
        blue_ver = blue_cands[0]['version'] if blue_cands else None

        if args.blue:
            if not blue_ver:
                print('❌ 无 BLUE 模型'); sys.exit(1)
            explain_model(df, blue_ver, f'评分卡BLUE  {code}  报价日 {issue_date}')
        else:
            explain_model(df, green_ver, f'评分卡GREEN(生产)  {code} {args.name}  报价日 {issue_date}')
            if blue_ver:
                explain_model(df, blue_ver, f'评分卡BLUE(对比)  {code}  报价日 {issue_date}')
    finally:
        try:
            os.remove(scored)
        except OSError:
            pass


if __name__ == '__main__':
    main()
