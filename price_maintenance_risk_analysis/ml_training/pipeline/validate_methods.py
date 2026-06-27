#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
多模型测试集验证：4 模型 × 2 测试集

模型：
  ① IV Top-N 评分卡 (WOE+LR)
  ② Lasso 选字段评分卡
  ③ 双向逐步回归选字段评分卡
  ④ LightGBM (清洗后全特征)

两套验证：
  Part A 内部 out-of-time：按报价日时序切 train=2020–2024 / test=2025–2026
  Part B 外部询转项目：全量训练 + 外部 Excel 测试

全程内存对象、不加载已存 pkl（规避 py3.7 sklearn 反序列化问题）。
评分卡测试集打分复用 predict_profitability.score_with_scorecard 的 WOE-apply 逻辑。

用法:
  python validate_methods.py data/features_derived.parquet \
      --external <外部测试Excel> [--threshold -10] [--n 12] [--iv-min 0.05] [--detail]
"""

import os
import sys
import argparse
import warnings
import numpy as np
import pandas as pd

warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
sys.path.insert(0, os.path.dirname(SCRIPT_DIR))   # ml_training/(compare_selection/predict_profitability 在根)

from train.train_scorecard import (
    calc_iv_all_features, woe_transform, train_scorecard_lr, build_scorecard,
)
from report.compare_selection import lasso_select, stepwise_aic
from predict_profitability import _parse_interval
from feature_exclusions import get_excluded_columns


# ====== 纯函数 ======

def calc_ks(y_true, y_score):
    """KS 统计量 (纯 numpy, 抄自 validate_model.py:210)。"""
    y_true = np.asarray(y_true)
    y_score = np.asarray(y_score)
    pos = y_score[y_true == 1]
    neg = y_score[y_true == 0]
    thresholds = np.sort(np.unique(y_score))
    if len(thresholds) < 2:
        return 0.0
    return max(abs((pos >= t).mean() - (neg >= t).mean()) for t in thresholds)


def eval_metrics(y_true, y_score):
    """返回 {auc, ks, n}。"""
    from sklearn.metrics import roc_auc_score
    y_true = pd.Series(y_true).reset_index(drop=True)
    y_score = pd.Series(y_score).reset_index(drop=True)
    valid = y_true.notna() & np.isfinite(y_score.values)
    yt = y_true[valid].astype(int).values
    ys = y_score[valid].values
    if len(yt) < 20 or len(set(yt.tolist())) < 2:
        return {'auc': float('nan'), 'ks': float('nan'), 'n': int(valid.sum())}
    return {'auc': float(roc_auc_score(yt, ys)),
            'ks': float(calc_ks(yt, ys)), 'n': int(valid.sum())}


def apply_woe_score(X, bins_dict, lr_model, features, base_points, B):
    """内存版 WOE 归箱+打分 (镜像 predict_profitability.score_with_scorecard line 71-100)。

    score = base_points + Σ(B · coef_i · woe_i)
    归箱 l<v<=r；v<=min_left→首箱；v>max_right→末箱；NaN/缺列→跳过(贡献0)。
    """
    coefs = dict(zip(features, lr_model.coef_[0]))
    scores = np.full(len(X), float(base_points), dtype=float)
    for feat in features:
        if feat not in X.columns or feat not in bins_dict:
            continue
        parsed = []
        for k, woe in bins_dict[feat].get('woe_map', {}).items():
            iv = _parse_interval(k)
            if iv is not None:
                parsed.append((iv[0], iv[1], float(woe)))
        if not parsed:
            continue
        parsed.sort(key=lambda t: t[0])
        min_left, first_woe, last_woe = parsed[0][0], parsed[0][2], parsed[-1][2]

        vals = pd.to_numeric(X[feat], errors='coerce').values
        coef = coefs.get(feat, 0.0)
        for i, v in enumerate(vals):
            if v != v:  # NaN → 跳过
                continue
            woe = None
            for l, r, w in parsed:
                if l < v <= r:
                    woe = w
                    break
            if woe is None:
                woe = first_woe if v <= min_left else last_woe
            scores[i] += B * coef * woe
    return scores


def score_to_proba(points, lr_model, base_points, B):
    """评分卡得分点 → 概率。logit = intercept + (score-base_points)/B。"""
    intercept = float(lr_model.intercept_[0])
    logit = intercept + (np.asarray(points, dtype=float) - base_points) / B
    return 1.0 / (1.0 + np.exp(-logit))


def decile_table(name, y_true, proba, ret):
    """按预测概率分十位，展示每段实际盈利率与平均7月收益(%)。"""
    y = pd.Series(y_true).reset_index(drop=True)
    p = pd.Series(proba).reset_index(drop=True)
    r = pd.Series(ret).reset_index(drop=True)
    valid = y.notna() & np.isfinite(p.values)
    d = pd.DataFrame({'y': y[valid].astype(int).values,
                      'p': p[valid].values, 'r': r[valid].values})
    if len(d) < 20:
        print(f'  [{name}] 样本不足, 跳过十分位')
        return
    d['q'] = pd.qcut(d['p'].rank(method='first'), 10, labels=[f'Q{i}' for i in range(1, 11)])
    print(f'\n  【{name}】十分位校准 (Q1=预测最低 → Q10=预测最高)')
    print(f'  {"分段":>5} {"n":>4} {"预测概率均":>10} {"概率范围":>18} {"实际盈利率":>9} {"平均7月收益":>11}')
    for q in [f'Q{i}' for i in range(1, 11)]:
        g = d[d['q'] == q]
        if len(g) == 0:
            continue
        print(f'  {q:>5} {len(g):>4} {g["p"].mean():>10.3f} [{g["p"].min():.3f}~{g["p"].max():.3f}]'
              f' {g["y"].mean()*100:>8.1f}% {g["r"].mean():>+10.2f}%')


# ====== 特征制备 + 选字段 ======

def make_features(df, threshold=-10, label_col=None, ret_col='7个月涨跌幅'):
    """数值特征(保留NaN不填充) + 标签 y + 连续收益 ret。丢弃标签NaN行(灰度剔除靠此)。

    label_col: 显式标签列(如 标签_盈利_-10_3m / 标签_极性_灰度剔除_7m); None→按 threshold 合成。
    ret_col:   连续收益列(默认 7个月涨跌幅), 用于灰度分桶; 不进 X。
    """
    if label_col is None:
        label_col = f'标签_盈利_{int(threshold)}'
        if label_col not in df.columns and ret_col in df.columns:
            df = df.copy()
            df[label_col] = (pd.to_numeric(df[ret_col], errors='coerce') > threshold / 100).astype(int)
    if label_col not in df.columns:
        print(f'❌ 找不到标签列 {label_col}')
        return None, None, None
    y = df[label_col]
    ret = pd.to_numeric(df[ret_col], errors='coerce') if ret_col in df.columns else pd.Series(np.nan, index=df.index)
    valid = y.notna()
    df = df.loc[valid].reset_index(drop=True)
    y = y.loc[valid].reset_index(drop=True)
    ret = ret.loc[valid].reset_index(drop=True)
    # 排除标签 + 统一剔除清单(含多期限原始收益列); ret 已单独取出不进 X
    exclude = set(get_excluded_columns(df.columns)) | {c for c in df.columns if '标签' in c}
    num_cols = [c for c in df.select_dtypes(include=[np.number]).columns if c not in exclude]
    X = df[num_cols].apply(lambda s: pd.to_numeric(s, errors='coerce')).reset_index(drop=True)
    return X, y, ret


def prep_features(X, medians=None):
    """套排除清单 + median 填充 + inf→0。medians=None 时用自身 median。"""
    excl = [c for c in get_excluded_columns(X.columns) if c in X.columns]
    X = X.drop(columns=excl)
    if medians is None:
        medians = X.median()
    X = X.fillna(medians).replace([np.inf, -np.inf], 0)
    # 仍可能有 NaN(median 为 NaN 的列) → 0
    X = X.fillna(0)
    return X, medians


def select_and_train_scorecards(X_train, y_train, iv_min, n):
    """在训练数据上 IV 池 + WOE + 三方法选字段 + 训练 LR。

    Returns:
        models: {方法名: dict(features, lr_model, base_points, B, auc_cv, auc_std, bins_dict)}
        bins_dict: 训练集 WOE 分箱(测试集归箱依据)
    """
    iv_df = calc_iv_all_features(X_train, y_train)
    pool = iv_df[iv_df['iv'] >= iv_min]['feature'].tolist()
    print(f'  IV>={iv_min} 候选池: {len(pool)} 个')
    X_woe_train, bins_dict = woe_transform(X_train, y_train, pool)
    woe_pool = [f for f in pool if f in X_woe_train.columns and X_woe_train[f].notna().any()]
    print(f'  WOE 变换成功: {len(woe_pool)} 个')

    selections = {
        f'IV Top-{n}': iv_df[iv_df['feature'].isin(woe_pool)].nlargest(n, 'iv')['feature'].tolist(),
        'Lasso': lasso_select(X_woe_train, y_train, woe_pool),
        '逐步回归': stepwise_aic(X_woe_train, y_train, woe_pool),
    }

    models = {}
    for name, feats in selections.items():
        if not feats:
            print(f'  [{name}] 无选中字段，跳过')
            continue
        lr_model, auc_cv, auc_std = train_scorecard_lr(X_woe_train, y_train, feats)
        _, base_points, B = build_scorecard(lr_model, bins_dict, feats)
        models[name] = {
            'features': feats, 'lr_model': lr_model,
            'base_points': base_points, 'B': B,
            'auc_cv': auc_cv, 'auc_std': auc_std, 'bins_dict': bins_dict,
        }
        print(f'  [{name}] {len(feats)}字段, train CV AUC={auc_cv:.3f}±{auc_std:.3f}')
    return models, bins_dict


def train_lgbm_clean(X, y):
    """清洗后全特征 LGBM (参数照搬 train_models.py:178-182)。"""
    import lightgbm as lgb
    from sklearn.model_selection import StratifiedKFold, cross_val_score
    model = lgb.LGBMClassifier(
        n_estimators=300, max_depth=5, learning_rate=0.03, num_leaves=31,
        is_unbalance=True, subsample=0.8, colsample_bytree=0.8,
        reg_alpha=0.1, reg_lambda=0.1, random_state=42, verbose=-1,
    )
    cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=42)
    auc = cross_val_score(model, X, y, cv=cv, scoring='roc_auc')
    model.fit(X, y)
    return model, float(auc.mean()), float(auc.std())


def score_lgbm(model, X_test, train_cols):
    """LGBM 打分：对齐到训练列(缺列补0)。"""
    Xt = X_test.reindex(columns=train_cols, fill_value=0).replace([np.inf, -np.inf], 0).fillna(0)
    return model.predict_proba(Xt)[:, 1]


def score_all(models, lgbm_model, lgbm_cols, X_test, y_test, train_cv_auc=None):
    """对 4 个模型在测试集上打分+评估。返回结果行列表。"""
    rows = []
    for name, m in models.items():
        scores = apply_woe_score(X_test, m['bins_dict'], m['lr_model'],
                                 m['features'], m['base_points'], m['B'])
        mt = eval_metrics(y_test, scores)
        row = {'模型': name, 'n_feat': len(m['features']),
               'train_cv_auc': m.get('auc_cv'), 'train_cv_std': m.get('auc_std'),
               'test_auc': mt['auc'], 'test_ks': mt['ks'], 'test_n': mt['n']}
        rows.append(row)
    # LGBM
    proba = score_lgbm(lgbm_model, X_test, lgbm_cols)
    mt = eval_metrics(y_test, proba)
    rows.append({
        '模型': 'LightGBM', 'n_feat': len(lgbm_cols),
        'train_cv_auc': train_cv_auc[0] if train_cv_auc else None,
        'train_cv_std': train_cv_auc[1] if train_cv_auc else None,
        'test_auc': mt['auc'], 'test_ks': mt['ks'], 'test_n': mt['n'],
    })
    return rows


def print_results(rows, title, coverage=None):
    print('\n' + '=' * 78)
    print(title)
    print('=' * 78)
    hdr = f'{"模型":<14}{"字段数":>6}{"train CV AUC":>16}{"test AUC":>12}{"test KS":>10}{"n":>6}'
    print(hdr)
    print('-' * 78)
    for r in rows:
        cv = f"{r['train_cv_auc']:.3f}±{r['train_cv_std']:.3f}" if r.get('train_cv_auc') else '-'
        print(f"{r['模型']:<14}{r['n_feat']:>6}{cv:>16}{r['test_auc']:>12.3f}{r['test_ks']:>10.3f}{r['test_n']:>6}")
    if coverage is not None:
        print(f'\n  ⚠ 特征覆盖率(填充前)~{coverage:.1f}%，大量 median 填充，结果仅供参考。')


# ====== Part A: 内部 OOT ======

def run_part_a(features_path, threshold, n, iv_min, split_year=2024,
               label_col=None, ret_col='7个月涨跌幅'):
    print('\n' + '#' * 78)
    print(f'# Part A — 内部 out-of-time 验证 (train ≤{split_year} / test ≥{split_year + 1})'
          + (f'  [label={label_col}]' if label_col else ''))
    print('#' * 78)

    df = pd.read_parquet(features_path)
    df = df.dropna(subset=['报价日']).copy()
    df['_year'] = (pd.to_numeric(df['报价日'], errors='coerce') // 10000).astype(int)
    df_train = df[df['_year'] <= split_year].drop(columns=['_year'])
    df_test = df[df['_year'] >= split_year + 1].drop(columns=['_year'])
    print(f'  时序切分: train {len(df_train)} 行, test {len(df_test)} 行')

    X_train_raw, y_train, _ = make_features(df_train, threshold, label_col=label_col, ret_col=ret_col)
    X_test_raw, y_test, _ = make_features(df_test, threshold, label_col=label_col, ret_col=ret_col)
    if X_train_raw is None or X_test_raw is None:
        return None
    X_test_raw = X_test_raw.reindex(columns=X_train_raw.columns)  # 对齐列

    X_train, medians = prep_features(X_train_raw)
    X_test, _ = prep_features(X_test_raw, medians=medians)  # 测试集用训练 median

    print(f'  train: {len(y_train)} 样本(正{y_train.mean()*100:.1f}%), '
          f'test: {len(y_test)} 样本(正{y_test.mean()*100:.1f}%) '
          f'⚠ 先验漂移→主看 AUC/KS 排序')

    print('\n  选字段+训练评分卡(只在 train 上):')
    models, _ = select_and_train_scorecards(X_train, y_train, iv_min, n)

    print('\n  训练 LightGBM:')
    lgbm_model, lgb_auc, lgb_std = train_lgbm_clean(X_train, y_train)
    print(f'  [LightGBM] {X_train.shape[1]}特征, train CV AUC={lgb_auc:.3f}±{lgb_std:.3f}')

    rows = score_all(models, lgbm_model, list(X_train.columns), X_test, y_test,
                     train_cv_auc=(lgb_auc, lgb_std))
    print_results(rows, 'Part A: 内部 OOT 验证 (test=2025–2026)')

    # 字段明细
    print('\n  各评分卡选中字段:')
    for name, m in models.items():
        print(f'    [{name}] {", ".join(m["features"])}')

    return rows


# ====== Part B: 外部询转 ======

def build_external_features(excel_path):
    """从 _scored Excel 重建完整特征矩阵(复用 validate_model.py 范本)。"""
    from export_features import load_scored_features, load_db_features, load_financial_ratios
    from derive_features import (
        derive_fcf_growth_rates, derive_fcf_cross_metrics,
        derive_financial_score_deltas, derive_valuation_relative,
        derive_market_momentum, derive_industry_valuation_growth,
        derive_market_index_features,
    )

    raw = pd.read_excel(excel_path, sheet_name='Sheet1')
    scored = load_scored_features(excel_path)

    # 报价日 → YYYYMMDD 串(load_scored_features 不带报价日, 从原始 Excel 取)
    # 报价日在 Excel 里是 int64 YYYYMMDD(如 20201202)，直接转串；to_datetime 会当纳秒误读
    rd = pd.to_numeric(raw['报价日'], errors='coerce')
    scored['报价日_excel'] = rd.apply(lambda x: str(int(x)) if pd.notna(x) else None)
    sample_keys = list(zip(scored['股票代码'].astype(str), scored['报价日_excel'].tolist()))

    print(f'  DB 特征提取({len(sample_keys)} 条)...')
    db_feats = load_db_features(sample_keys)
    db_cols = [c for c in db_feats.columns if c != '股票代码']
    dup = [c for c in db_cols if c in scored.columns]
    if dup:
        db_feats = db_feats.drop(columns=dup)
        db_cols = [c for c in db_feats.columns if c != '股票代码']
    df = pd.concat([scored.reset_index(drop=True),
                    db_feats[db_cols].reset_index(drop=True)], axis=1)

    # 财务比率(financial_indicators) PIT —— 与训练 main() 同源(按报价日回溯)
    try:
        sk = list(zip(scored['股票代码'].astype(str), scored['报价日_excel'].astype(str)))
        ratio_feats = load_financial_ratios(sk)
        if ratio_feats is not None and not ratio_feats.empty:
            rcols = [c for c in ratio_feats.columns if c not in df.columns]
            if rcols:
                df = pd.concat([df.reset_index(drop=True), ratio_feats[rcols].reset_index(drop=True)], axis=1)
                print(f'  财务比率 PIT(financial_indicators): +{len(rcols)} 列')
    except Exception as e:
        print(f'  财务比率跳过: {e}')

    # 类型清理
    str_keep = {'股票代码', '股票简称', '一级行业', '二级行业', '三级行业',
                '最终结论', '定增决策', '行业代码', '行业名称'}
    for c in df.columns:
        if c in str_keep:
            df[c] = df[c].astype(str)
        elif df[c].dtype == object or str(df[c].dtype) == 'category':
            df[c] = pd.to_numeric(df[c], errors='coerce')

    for f in [derive_fcf_growth_rates, derive_fcf_cross_metrics,
              derive_financial_score_deltas, derive_valuation_relative,
              derive_market_momentum]:
        df = f(df)
    try:
        df = derive_industry_valuation_growth(df)
        df = derive_market_index_features(df)
    except Exception as e:
        print(f'  DB 衍生跳过: {e}')
    df = df.replace([np.inf, -np.inf], np.nan)
    return df


def run_part_b(features_path, external_excel, threshold, n, iv_min, detail=False):
    print('\n' + '#' * 78)
    print('# Part B — 外部询转验证 (训练集剔除外部样本后训练 + 外部 401 测试)')
    print('#' * 78)

    # 外部样本键(股票代码+报价日) —— 从训练集剔除，保证测试集与训练集不重叠
    raw_ext = pd.read_excel(external_excel, sheet_name='Sheet1')
    ext_keys = set(zip(
        raw_ext['股票代码'].astype(str),
        pd.to_numeric(raw_ext['报价日'], errors='coerce')
            .apply(lambda x: str(int(x)) if pd.notna(x) else '')))

    # 全量训练数据，先剔除外部样本
    df_full = pd.read_parquet(features_path)
    tr_keys = list(zip(
        df_full['股票代码'].astype(str),
        pd.to_numeric(df_full['报价日'], errors='coerce')
            .apply(lambda x: str(int(x)) if pd.notna(x) else '')))
    keep = [k not in ext_keys for k in tr_keys]
    n_before = len(df_full)
    df_full = df_full[keep].reset_index(drop=True)
    print(f'  训练集剔除外部 {len(ext_keys)} 键: {n_before} → {len(df_full)} (保证无重叠)')

    X_full_raw, y_full, _ = make_features(df_full, threshold)
    if X_full_raw is None:
        return None
    X_full, medians = prep_features(X_full_raw)
    print(f'  训练: {len(y_full)} 样本(正{y_full.mean()*100:.1f}%), {X_full.shape[1]} 特征')

    print('\n  选字段+训练评分卡(在全量上):')
    models, _ = select_and_train_scorecards(X_full, y_full, iv_min, n)

    print('\n  训练 LightGBM:')
    lgbm_model, lgb_auc, lgb_std = train_lgbm_clean(X_full, y_full)
    print(f'  [LightGBM] {X_full.shape[1]}特征, train CV AUC={lgb_auc:.3f}±{lgb_std:.3f}')

    # 外部测试特征
    print('\n  构建外部测试特征:')
    df_ext = build_external_features(external_excel)
    X_ext_raw, y_ext, ret_ext = make_features(df_ext, threshold)
    if X_ext_raw is None:
        print('  ⚠ 外部集无标签列, Part B 中止')
        return None
    X_ext_raw = X_ext_raw.reindex(columns=X_full.columns)  # 对齐到全量列
    X_ext, _ = prep_features(X_ext_raw, medians=medians)   # 用全量 median 填充
    print(f'  外部测试: {len(y_ext)} 样本(正{y_ext.mean()*100:.1f}%, 标签非空{y_ext.notna().sum()})')

    # 覆盖率(填充前)
    cov_cells = X_ext_raw.notna().sum().sum()
    cov_total = X_ext_raw.shape[0] * X_ext_raw.shape[1]
    coverage = cov_cells / cov_total * 100 if cov_total else 0.0

    rows = score_all(models, lgbm_model, list(X_full.columns), X_ext, y_ext,
                     train_cv_auc=(lgb_auc, lgb_std))
    print_results(rows, 'Part B: 外部询转验证 (401 条)', coverage=coverage)

    if detail:
        print('\n  ===== 十分位区分度校准 (外部 401) =====')
        for name, m in models.items():
            pts = apply_woe_score(X_ext, m['bins_dict'], m['lr_model'],
                                  m['features'], m['base_points'], m['B'])
            proba = score_to_proba(pts, m['lr_model'], m['base_points'], m['B'])
            decile_table(name, y_ext, proba, ret_ext)
        proba_lgb = score_lgbm(lgbm_model, X_ext, list(X_full.columns))
        decile_table('LightGBM', y_ext, proba_lgb, ret_ext)
    return rows


# ====== main ======

def main():
    ap = argparse.ArgumentParser(description='多模型测试集验证')
    ap.add_argument('features_path', help='features_derived.parquet')
    ap.add_argument('--external', help='外部测试 Excel(_scored 格式)')
    ap.add_argument('--threshold', type=float, default=-10)
    ap.add_argument('--n', type=int, default=12, help='IV Top-N')
    ap.add_argument('--iv-min', type=float, default=0.05)
    ap.add_argument('--detail', action='store_true', help='打印混淆矩阵/分位校准(暂留)')
    ap.add_argument('--only', choices=['a', 'b'], help='只跑 Part A 或 B')
    args = ap.parse_args()

    rows_a = rows_b = None
    if args.only != 'b':
        rows_a = run_part_a(args.features_path, args.threshold, args.n, args.iv_min)
    if args.only != 'a' and args.external:
        rows_b = run_part_b(args.features_path, args.external, args.threshold, args.n, args.iv_min, detail=args.detail)
    elif args.only != 'a' and not args.external:
        print('\n(未提供 --external, 跳过 Part B)')

    # 落 CSV
    out_dir = os.path.join(os.path.dirname(SCRIPT_DIR), 'output')  # ml_training/output
    os.makedirs(out_dir, exist_ok=True)
    if rows_a:
        pd.DataFrame(rows_a).to_csv(
            os.path.join(out_dir, 'validation_partA_oot.csv'),
            index=False, encoding='utf-8-sig')
        print(f"\n✅ Part A 已保存: output/validation_partA_oot.csv")
    if rows_b:
        pd.DataFrame(rows_b).to_csv(
            os.path.join(out_dir, 'validation_partB_external.csv'),
            index=False, encoding='utf-8-sig')
        print(f"✅ Part B 已保存: output/validation_partB_external.csv")


if __name__ == '__main__':
    main()
