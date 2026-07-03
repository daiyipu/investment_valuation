#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""评分卡审计报告 —— 上线前人工核查基础。

为何: 评分卡上线前审核人员需看到每个指标的 业务解释/IV/贡献度/分箱/分值,
      不能只看一个 AUC。本脚本把 current.full(或指定版本)的 SC 评分卡
      全部明细导出存档(CSV + Markdown)。

每特征输出:
  业务解释 | IV | LR系数 | 贡献度(logit摆幅=coef×(maxWOE-minWOE)) | 贡献占比% |
  分箱边界 | 各箱WOE | 各箱分值(=coef×WOE×100)
  单调性检查结果(新增)

数据来源: lr_bundle(woe_bins/lr_model/proba_deciles) + features_derived 重算 IV。
业务解释: feature_glossary.explain()。

新增功能:
  - 支持 scorecardpy 树方法分箱 (--method tree)
  - WOE 单调性检查和自动调整
  - 手动分箱合并方案 (--manual-merge)
  - 质量评估和优化建议

用法:
  python report/audit_scorecard.py                    # 审计 current.full
  python report/audit_scorecard.py <version>
  python report/audit_scorecard.py --method tree     # 使用树方法分箱
  python report/audit_scorecard.py --manual-merge    # 使用手动合并方案
"""
import os
import sys
import pickle
import argparse

import numpy as np
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # ml_training/
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'pipeline'))   # 管线模块已移入 pipeline/
from deploy.model_registry import get_current
from deploy.db_model_store import load_predict_bundle, get_model_meta
from train.train_scorecard import calc_iv_all_features
from report.feature_glossary import explain

# 尝试导入scorecardpy
try:
    import scorecardpy as sc
    SCORECARDPY_AVAILABLE = True
except ImportError:
    SCORECARDPY_AVAILABLE = False
    print("⚠️ scorecardpy 未安装，将使用传统方法")

PARQUET = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'features_derived.parquet')


def _gray_label(df, horizon=7):
    """horizon 个月 gray 标签, 用生产口径 GRAY_CFG(与训练一致)。"""
    from train.train_horizon_models import GRAY_CFG
    lo, hi = GRAY_CFG.get(horizon, (-20, 10))
    col = f'{horizon}个月涨跌幅'
    s = pd.to_numeric(df[col], errors='coerce') if col in df.columns else pd.Series(np.nan, index=df.index)
    return pd.Series(np.where(s > hi, 1, np.where(s < lo, 0, np.nan)), index=df.index)


def check_woe_monotonicity(bins, features):
    """检查WOE单调性 - 识别波浪形特征"""
    if isinstance(bins, dict):
        # 传统格式：逐个特征检查
        monotonicity_results = {}
        for feature in features:
            wb = bins.get(feature, {})
            woes = wb.get('woes', [])

            if len(woes) < 2:
                monotonicity_results[feature] = {
                    'status': 'insufficient_bins',
                    'description': '分箱数量不足',
                    'is_acceptable': False
                }
                continue

            # 检查单调性
            trends = []
            for i in range(1, len(woes)):
                if woes[i] >= woes[i-1]:
                    trends.append('increase')
                else:
                    trends.append('decrease')

            unique_trends = len(set(trends))

            if unique_trends == 1:
                if trends[0] == 'increase':
                    status = 'monotonic_increase'
                    description = '单调递增'
                else:
                    status = 'monotonic_decrease'
                    description = '单调递减'
            elif unique_trends == 2:
                status = 'u_shape'
                description = 'U型'
            else:
                status = 'wave'
                description = f'波浪形({unique_trends-1}次趋势变化)'

            monotonicity_results[feature] = {
                'status': status,
                'description': description,
                'is_acceptable': unique_trends <= 2,
                'unique_trends': unique_trends,
                'bin_count': len(woes)
            }

        return monotonicity_results

    else:
        # scorecardpy格式
        if isinstance(bins, dict):
            bins_df = pd.concat(bins, ignore_index=True)
        else:
            bins_df = bins

        monotonicity_results = {}
        for feature in features:
            feat_bins = bins_df[bins_df['variable'] == feature].copy()

            if len(feat_bins) < 2:
                monotonicity_results[feature] = {
                    'status': 'insufficient_bins',
                    'description': '分箱数量不足',
                    'is_acceptable': False
                }
                continue

            try:
                feat_bins = feat_bins.sort_values('bin')
                feat_bins['badprob2'] = feat_bins['badprob'].shift(1)
                feat_bins = feat_bins.dropna(subset=['badprob2'])
                feat_bins['badprob_trend'] = feat_bins['badprob'] >= feat_bins['badprob2']

                unique_trends = feat_bins['badprob_trend'].nunique()

                if unique_trends == 1:
                    if feat_bins['badprob_trend'].all():
                        status = 'monotonic_increase'
                        description = '单调递增'
                    else:
                        status = 'monotonic_decrease'
                        description = '单调递减'
                elif unique_trends == 2:
                    status = 'u_shape'
                    description = 'U型'
                else:
                    status = 'wave'
                    description = f'波浪形({unique_trends-1}次趋势变化)'

                monotonicity_results[feature] = {
                    'status': status,
                    'description': description,
                    'unique_trends': unique_trends,
                    'is_acceptable': unique_trends <= 2,
                    'bin_count': len(feat_bins)
                }
            except Exception as e:
                monotonicity_results[feature] = {
                    'status': 'error',
                    'description': f'检查失败: {e}',
                    'is_acceptable': False,
                    'bin_count': len(feat_bins) if 'feat_bins' in locals() else 0
                }

        return monotonicity_results


def enhanced_woe_binning_scorecardpy(X, y, features, method="tree"):
    """使用scorecardpy进行树方法分箱和单调性优化"""
    if not SCORECARDPY_AVAILABLE:
        print("  ⚠️ scorecardpy不可用，使用传统分箱方法")
        return None

    try:
        from validate.eval_loyo import fit_woe

        # 准备数据 - 确保索引对齐
        dt = X.copy().reset_index(drop=True)
        y_aligned = y.reset_index(drop=True) if hasattr(y, 'reset_index') else pd.Series(y)

        # 确保y是数值型二分类标签（0和1）
        y_numeric = y_aligned.astype(int)

        # 检查是否真的是二分类
        unique_vals = y_numeric.unique()
        if len(unique_vals) < 2:
            print(f"  ❌ y列只有一个唯一值: {unique_vals}，无法分箱")
            return None

        dt['y'] = y_numeric.values

        print(f"  数据准备完成: {dt.shape}, y分布: {pd.Series(y_numeric).value_counts().to_dict()}")

        # 1. scorecardpy树方法分箱
        print(f"  使用scorecardpy {method}方法分箱...")
        bins = sc.woebin(
            dt, 'y', features,
            method=method,
            count_distr_limit=0.05,
            bin_num_limit=8,
            stop_limit=0.1,
            print_step=0
        )

        # 2. 单调性检查
        print("  检查WOE单调性...")
        monotonicity_results = check_woe_monotonicity(bins, features)

        # 3. 自动调整有问题的分箱
        problematic = [f for f, r in monotonicity_results.items()
                     if not r['is_acceptable']]
        if problematic:
            print(f"  发现{len(problematic)}个波浪形特征，尝试调整...")
            try:
                bins_adj = sc.woebin_adj(
                    dt, 'y', bins,
                    adj_all_var=False,
                    method=method,
                    save_breaks_list=None,
                    count_distr_limit=0.05
                )
                bins = bins_adj
                monotonicity_results = check_woe_monotonicity(bins, features)
                print(f"  ✅ 分箱调整完成")
            except Exception as e:
                print(f"  ⚠️ 自动调整失败: {e}")
                print("  使用原始分箱继续")

        return bins, monotonicity_results

    except Exception as e:
        print(f"  ❌ scorecardpy处理失败: {e}")
        return None


def manual_bin_merging(X, y, features):
    """手动分箱合并方案 - 针对波浪形特征"""
    print("  应用手动分箱合并方案...")

    bins = {}

    for feature in features:
        feat_data = X[feature].copy()

        if feature == 'sue_zscore':
            # 5个分箱 → 2个分箱
            # (-25298258.205, 0.148] → 高风险区间
            # (0.148, 1446047487.367] → 低风险区间
            bins_cat = pd.cut(feat_data, bins=[-np.inf, 0.148, np.inf], labels=False)
            bins[feature] = {'rights': [0.148], 'woes': [], 'bins': bins_cat}

        elif feature == 'mf_main_net_ratio_5d':
            # 5个分箱 → 3个分箱
            # (-0.175, -0.00764] → 高风险
            # (-0.00764, 0.00584] → 中风险
            # (0.00584, 0.339] → 低风险
            bins_cat = pd.cut(feat_data, bins=[-np.inf, -0.00764, 0.00584, np.inf], labels=False)
            bins[feature] = {'rights': [-0.00764, 0.00584], 'woes': [], 'bins': bins_cat}

        else:
            # 其他特征使用原始5分箱
            from validate.eval_loyo import fit_woe
            _, original_bins = fit_woe(X, y, [feature])
            original_wb = original_bins.get(feature, {})
            bins[feature] = {
                'rights': original_wb.get('rights', []),
                'woes': original_wb.get('woes', []),
                'bins': None  # 标记为未修改
            }

    # 计算合并特征的WOE值
    for feature in ['sue_zscore', 'mf_main_net_ratio_5d']:
        if feature in bins and bins[feature]['bins'] is not None:
            feat_bins = bins[feature]['bins']
            y_valid = y[feat_bins.notna()]
            bins_valid = feat_bins[feat_bins.notna()]

            woes = []
            for bin_val in sorted(bins_valid.unique()):
                mask = bins_valid == bin_val
                goods = ((y_valid == 1) & mask).sum()
                bads = ((y_valid == 0) & mask).sum()
                total_goods = (y_valid == 1).sum()
                total_bads = (y_valid == 0).sum()

                if goods == 0 or bads == 0:
                    woe = 0
                else:
                    dist_good = goods / total_goods
                    dist_bad = bads / total_bads
                    woe = np.log(dist_good / dist_bad) if dist_bad > 0 else 0

                woes.append(woe)

            bins[feature]['woes'] = woes
            del bins[feature]['bins']  # 清理临时数据

    return bins


def audit(version, out_dir, method='original', manual_merge=False):
    """
    生成评分卡审计报告

    Args:
        version: 模型版本
        out_dir: 输出目录
        method: 分箱方法 ('original', 'tree', 'chimerge')
        manual_merge: 是否使用手动合并方案
    """
    bundle = load_predict_bundle(version)
    sc = pickle.loads(bundle['lr_bundle'])
    feats = sc['features']
    woe_bins = sc['woe_bins']
    lr = sc['lr_model']
    coefs = dict(zip(feats, lr.coef_[0]))
    intercept = float(lr.intercept_[0])
    meta = get_model_meta(version) or {}
    horizon = meta.get('horizon', 7)

    # 重算 IV(训练集, 同期限灰度标签)
    df = pd.read_parquet(PARQUET)
    y = _gray_label(df, horizon)
    X = df[feats].apply(pd.to_numeric, errors='coerce')
    iv_tbl = calc_iv_all_features(X, y)
    iv_map = dict(zip(iv_tbl['feature'], iv_tbl['iv']))

    # 根据方法选择分箱方式
    monotonicity_results = None
    enhanced_bins = None

    if method == 'original' and not manual_merge:
        # 原始方法
        print("使用原始等频分箱方法")
        final_bins = woe_bins
    elif manual_merge:
        # 手动合并方案
        print("使用手动分箱合并方案")
        enhanced_bins = manual_bin_merging(X, y, feats)
        final_bins = enhanced_bins
    elif method in ['tree', 'chimerge']:
        # scorecardpy树方法
        print(f"使用scorecardpy {method}方法")
        result = enhanced_woe_binning_scorecardpy(X, y, feats, method)
        if result:
            enhanced_bins, monotonicity_results = result
            final_bins = enhanced_bins
        else:
            print("  回退到原始方法")
            final_bins = woe_bins
    else:
        final_bins = woe_bins

    # 单调性检查
    if monotonicity_results is None:
        monotonicity_results = check_woe_monotonicity(final_bins, feats)

    rows = []
    for f in feats:
        wb = final_bins.get(f, {})
        rights = list(wb.get('rights', []))
        woes = [float(w) for w in wb.get('woes', [])]
        coef = float(coefs.get(f, 0.0))
        pts = [round(coef * w * 100, 2) for w in woes]          # 各箱分值 = coef×WOE×100
        swing = float(coef * (max(woes) - min(woes))) if woes else 0.0

        # 获取单调性信息
        mono_info = monotonicity_results.get(f, {})
        mono_status = mono_info.get('status', 'unknown')
        mono_desc = mono_info.get('description', 'unknown')
        mono_acceptable = mono_info.get('is_acceptable', True)

        rows.append({
            '特征': f, '业务解释': explain(f),
            'IV': round(float(iv_map.get(f, np.nan)), 4),
            'LR系数': round(coef, 4),
            '贡献度_logit摆幅': round(swing, 4),
            '分箱边界': [round(r, 4) for r in rights],
            '各箱WOE': [round(w, 3) for w in woes],
            '各箱分值_x100': pts,
            '单调性': mono_desc,
            '单调性状态': mono_status,
            '单调性合格': mono_acceptable,
        })
    adf = pd.DataFrame(rows)
    tot = adf['贡献度_logit摆幅'].abs().sum() or 1.0
    adf['贡献度占比%'] = (adf['贡献度_logit摆幅'].abs() / tot * 100).round(1)
    adf = adf.reindex(adf['贡献度_logit摆幅'].abs().sort_values(ascending=False).index)

    os.makedirs(out_dir, exist_ok=True)
    csv_p = os.path.join(out_dir, 'scorecard_audit.csv')
    adf.to_csv(csv_p, index=False, encoding='utf-8-sig')

    dec = [round(float(x), 3) for x in sc.get('proba_deciles', [])]

    # 单调性统计
    total_features = len(monotonicity_results)
    acceptable_features = sum(1 for r in monotonicity_results.values() if r.get('is_acceptable', False))
    wave_features = sum(1 for r in monotonicity_results.values() if r.get('status') == 'wave')

    method_desc = {
        'original': '原始等频分箱',
        'tree': 'scorecardpy树方法',
        'chimerge': 'scorecardpy ChiMerge方法'
    }

    md = [f'# 评分卡审计报告 — {version}', '',
          f'## 基本信息',
          f'- 期限: {horizon}m gray | 特征数: {len(feats)} | 截距(intercept): {intercept:+.4f}',
          f'- 分箱方法: {method_desc.get(method, "未知")}',
          f'- 模型 LOYO: AUC {meta.get("sc_loyo_auc")} | KS {meta.get("sc_loyo_ks")}'
          f'  (若空则仅有 OOT: AUC {meta.get("sc_oot_auc")} / KS {meta.get("sc_oot_ks")})',
          f'- 档位边界(proba 10 分位): {dec}', '',
          f'## WOE单调性分析',
          f'- 总特征数: {total_features}',
          f'- 合格特征: {acceptable_features} ({acceptable_features/total_features*100:.1f}%)',
          f'- 波浪形特征: {wave_features}',
          f'- 质量评估: {"优秀" if acceptable_features/total_features >= 0.8 else "良好" if acceptable_features/total_features >= 0.6 else "需改进"}', '',


          '> **说明**:',
          '> 分值 = LR系数 × 该箱WOE × 100(logit 贡献量, 越大越推高盈利概率);',
          '> 贡献度 = 系数×(最高箱WOE−最低箱WOE), 即该特征能造成的 logit 最大摆幅;',
          '> 贡献占比 = 该特征|摆幅|占全部特征之和的比例;',
          '> 单调性: 单调递增/递减为✅合格，U型尚可，波浪形❌不合格。', '',

          f'## 特征详细审计',
          '| 特征 | 业务解释 | IV | LR系数 | 贡献度(摆幅) | 占比% | 分箱边界 | 各箱WOE | 各箱分值×100 | 单调性 | 合格 |',
          '|---|---|---|---|---|---|---|---|---|---|---|']
    for _, r in adf.iterrows():
        mono_emoji = '✅' if r['单调性合格'] else '❌'
        md.append(f'| {r["特征"]} | {r["业务解释"]} | {r["IV"]} | {r["LR系数"]:+.4f} | '
                  f'{r["贡献度_logit摆幅"]:+.4f} | {r["贡献度占比%"]:.1f} | '
                  f'{r["分箱边界"]} | {r["各箱WOE"]} | {r["各箱分值_x100"]} | '
                  f'{r["单调性"]} | {mono_emoji} |')

    # 添加优化建议
    md.append('')
    md.append('## 优化建议')
    if wave_features > 0:
        md.append(f'### 发现{wave_features}个波浪形特征需要优化：')
        for feature, result in monotonicity_results.items():
            if result.get('status') == 'wave':
                md.append(f'- **{feature}**: {result.get("description")} - 建议合并分箱')
        md.append('')
        md.append('**修正方案**:')
        md.append('1. 使用 `--method tree` 重新分箱（推荐）')
        md.append('2. 使用 `--manual-merge` 手动合并')
        md.append('3. 删除波浪形特征后重新训练')
    else:
        md.append('✅ 所有特征单调性检查通过，无需优化')

    md.append('')
    md.append('---')
    md.append('*本报告由 audit_scorecard.py 自动生成*')

    md_p = os.path.join(out_dir, 'scorecard_audit.md')
    with open(md_p, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(md))
    print(f'✅ 审计报告已存档: {out_dir}/  (scorecard_audit.csv + scorecard_audit.md)\n')

    # 显示结果
    show = adf[['特征', '业务解释', 'IV', 'LR系数', '贡献度_logit摆幅', '贡献度占比%', '单调性', '单调性合格']].copy()
    print(show.to_string(index=False))

    # 显示单调性统计
    print(f'\n单调性统计: {acceptable_features}/{total_features} 合格 ({acceptable_features/total_features*100:.1f}%)')
    if wave_features > 0:
        print(f'⚠️ 发现{wave_features}个波浪形特征需要优化')
    else:
        print('✅ 所有特征单调性检查通过')

    return adf


def main():
    ap = argparse.ArgumentParser(description='评分卡审计报告(业务解释+IV+贡献度+分箱+分值+单调性)')
    ap.add_argument('version', nargs='?', help='模型版本(默认 current.full)')
    ap.add_argument('--out', help='输出目录(默认 output/audit_<version>)')
    ap.add_argument('--method', choices=['original', 'tree', 'chimerge'], default='original',
                   help='分箱方法: original=原始等频分箱, tree/chimerge=scorecardpy树方法')
    ap.add_argument('--manual-merge', action='store_true',
                   help='使用手动分箱合并方案（针对波浪形特征）')
    args = ap.parse_args()
    version = args.version or get_current('full')
    out_dir = args.out or os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       'output', f'audit_{version}')

    print(f"=== 评分卡审计报告 ===")
    print(f"版本: {version}")
    print(f"分箱方法: {args.method}")
    if args.manual_merge:
        print(f"手动合并: 启用")
    print()

    audit(version, out_dir, method=args.method, manual_merge=args.manual_merge)


if __name__ == '__main__':
    main()
