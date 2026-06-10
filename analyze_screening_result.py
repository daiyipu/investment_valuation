#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增筛选结果透视分析脚本

用法:
  python analyze_screening_result.py <scored_excel_path>

输出:
  - 控制台汇总统计
  - <原文件名>_analysis.xlsx (多sheet透视表)
"""

import sys
import os
import pandas as pd
import numpy as np


def analyze(filepath):
    df = pd.read_excel(filepath, sheet_name='Sheet1')
    print(f'共 {len(df)} 只股票\n')

    # 预处理：涨跌幅转数值(去掉%号)
    for col in ['7个月后涨跌幅', '溢价率下限', '溢价率上限']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace('%', '').str.replace('+', ''),
                                    errors='coerce')

    output_path = filepath.replace('.xlsx', '_analysis.xlsx')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

        # ========== 1. 最终结论汇总 ==========
        print('='*60)
        print('1. 最终结论汇总')
        print('='*60)
        conclusion = df['最终结论'].value_counts()
        print(conclusion.to_string())
        print(f'\n通过率: {conclusion.get("通过", 0) / len(df) * 100:.1f}%')
        conclusion.to_frame('数量').to_excel(writer, sheet_name='1_最终结论汇总')

        # ========== 2. 定增决策汇总 ==========
        print(f'\n{"="*60}')
        print('2. 定增决策汇总')
        print('='*60)
        decision = df['定增决策'].value_counts()
        print(decision.to_string())
        decision.to_frame('数量').to_excel(writer, sheet_name='2_定增决策汇总')

        # ========== 3. 按一级行业透视 ==========
        print(f'\n{"="*60}')
        print('3. 按一级行业透视')
        print('='*60)
        industry_pivot = df.pivot_table(
            index='一级行业',
            values=['股票代码', '7个月后涨跌幅'],
            aggfunc={'股票代码': 'count', '7个月后涨跌幅': 'mean'},
        ).rename(columns={'股票代码': '数量', '7个月后涨跌幅': '平均7月涨幅'})
        industry_pivot = industry_pivot.sort_values('数量', ascending=False)
        # 加通过率
        pass_by_ind = df[df['最终结论'] == '通过'].groupby('一级行业').size()
        industry_pivot['通过数'] = pass_by_ind
        industry_pivot['通过率'] = (industry_pivot['通过数'] / industry_pivot['数量'] * 100).round(1)
        print(industry_pivot.head(15).to_string())
        industry_pivot.to_excel(writer, sheet_name='3_行业透视')

        # ========== 4. 子场景通过率 ==========
        print(f'\n{"="*60}')
        print('4. 子场景通过率')
        print('='*60)
        # 新版列名(市场指数/行业PE等)或旧版(①②③)
        sub_cols = [c for c in df.columns if c in
                    ['市场指数', '行业PE', '个股PE', 'DCF估值', '修正PE估值',
                     '参数构造', '蒙特卡洛', '反向推算',
                     '①历史数据场景', '②预期估值场景', '③其他场景']]
        if sub_cols:
            sub_stats = pd.DataFrame({
                '通过数': df[sub_cols].apply(lambda c: (c == '✓').sum() if c.dtype == object else (c == '✓ 通过').sum()),
                '通过率%': df[sub_cols].apply(lambda c: ((c == '✓').sum() if c.dtype == object else (c == '✓ 通过').sum()) / len(df) * 100).round(1)
            })
            print(sub_stats.to_string())
            sub_stats.to_excel(writer, sheet_name='4_子场景通过率')

            # 4b. 各子场景通过/不通过的盈利占比(盈利=涨跌幅>-10%)
            WIN_THRESHOLD = -0.10  # 盈利定义：7个月涨跌幅 > -10%
            print(f'\n  ── 各子场景盈亏占比(盈利定义: 涨跌幅>-10%) ──')
            ret_col_tmp = '7个月后涨跌幅'
            if ret_col_tmp in df.columns:
                df_pl = df.copy()
                df_pl[ret_col_tmp] = pd.to_numeric(df_pl[ret_col_tmp].astype(str).str.replace('%', '').str.replace('+', ''), errors='coerce')
                pl_rows = []
                for col in sub_cols:
                    is_pass = df_pl[col].astype(str).str.contains('✓')
                    for status, mask in [('通过', is_pass), ('不通过', ~is_pass)]:
                        sub = df_pl[mask][ret_col_tmp].dropna()
                        n = len(sub)
                        if n == 0:
                            continue
                        win = (sub > WIN_THRESHOLD).sum()
                        pl_rows.append({
                            '子场景': col, '状态': status, '总数': n,
                            '盈利数': win, '盈利占比%': round(win / n * 100, 1),
                            '平均收益%': round(sub.mean() * 100, 2),
                            '中位数%': round(sub.median() * 100, 2),
                        })
                pl_by_sub = pd.DataFrame(pl_rows)
                print(pl_by_sub.to_string(index=False))
                pl_by_sub.to_excel(writer, sheet_name='4b_子场景盈亏', index=False)

        # ========== 5. 财务评分统计 ==========
        print(f'\n{"="*60}')
        print('5. 财务评分统计(最新年)')
        print('='*60)
        score_cols = [c for c in df.columns if c.startswith('总分_') and c[3:].isdigit()]
        if score_cols:
            latest_year = max(c.split('_')[1] for c in score_cols)
            latest_col = f'总分_{latest_year}'
            rating_col = f'评级_{latest_year}'
            print(f'最新评分年份: {latest_year}')
            rating_dist = df[rating_col].value_counts() if rating_col in df.columns else pd.Series()
            print(rating_dist.to_string())
            score_stats = df[[latest_col]].describe()
            print(f'\n总分统计:\n{score_stats.to_string()}')
            rating_dist.to_frame('数量').to_excel(writer, sheet_name='5_财务评分统计')
            # 评分趋势
            trend_cols = [c for c in ['总分_趋势', '盈利能力_趋势', '成长能力_趋势', '综合趋势'] if c in df.columns]
            if trend_cols:
                trend_dist = df[trend_cols].apply(lambda c: c.value_counts())
                print(f'\n趋势分布:\n{trend_dist.to_string()}')
                trend_dist.to_excel(writer, sheet_name='5_财务评分统计', startrow=len(rating_dist)+4)

        # ========== 6. 7个月涨跌幅分析 ==========
        print(f'\n{"="*60}')
        print('6. 7个月后涨跌幅分析(解禁收益)')
        print('='*60)
        ret_col = '7个月后涨跌幅'
        if ret_col in df.columns:
            # 转数值
            ret = pd.to_numeric(df[ret_col], errors='coerce').dropna()
            print(f'有效数据: {len(ret)}只')
            print(f'平均涨跌: {ret.mean()*100:.2f}%')
            print(f'中位数: {ret.median()*100:.2f}%')
            print(f'盈利占比: {(ret > 0).sum()}/{len(ret)} = {(ret > 0).mean()*100:.1f}%')
            print(f'平均涨跌(>0): {ret[ret > 0].mean()*100:.2f}%' if (ret > 0).any() else '')
            print(f'平均涨跌(<0): {ret[ret < 0].mean()*100:.2f}%' if (ret < 0).any() else '')
            # 按结论分组
            df_ret = df.copy()
            df_ret[ret_col] = pd.to_numeric(df_ret[ret_col], errors='coerce')
            ret_by_conclusion = df_ret.groupby('最终结论')[ret_col].agg(['count', 'mean', 'median'])
            ret_by_conclusion['mean'] = (ret_by_conclusion['mean'] * 100).round(2)
            ret_by_conclusion['median'] = (ret_by_conclusion['median'] * 100).round(2)
            print(f'\n按结论分组:\n{ret_by_conclusion.to_string()}')
            ret_by_conclusion.to_excel(writer, sheet_name='6_7月涨跌幅分析')

            # 盈亏对比(按通过/不通过，盈利=涨跌幅>-10%)
            print(f'\n  ── 盈亏占比对比(盈利定义: 涨跌幅>-10%) ──')
            pl_rows = []
            for conclusion in ['通过', '不通过']:
                sub = df_ret[df_ret['最终结论'] == conclusion][ret_col].dropna()
                n = len(sub)
                if n == 0:
                    continue
                win = (sub > WIN_THRESHOLD).sum()
                lose = (sub <= WIN_THRESHOLD).sum()
                big_win = (sub > 0.3).sum()      # 大涨>30%
                big_lose = (sub < -0.3).sum()     # 大跌<-30%
                pl_rows.append({
                    '结论': conclusion,
                    '总数': n,
                    '盈利数': win,
                    '亏损数': lose,
                    '盈利占比%': round(win / n * 100, 1),
                    '亏损占比%': round(lose / n * 100, 1),
                    '大涨>30%占比': round(big_win / n * 100, 1),
                    '大跌<-30%占比': round(big_lose / n * 100, 1),
                    '平均收益%': round(sub.mean() * 100, 2),
                    '中位数收益%': round(sub.median() * 100, 2),
                })
            pl_df = pd.DataFrame(pl_rows).set_index('结论')
            print(pl_df.to_string())
            pl_df.to_excel(writer, sheet_name='6_7月涨跌幅分析', startrow=len(ret_by_conclusion) + 4)

        # ========== 6b. 收益率区间占比(按通过/不通过) ==========
        print(f'\n{"="*60}')
        print('6b. 收益率区间占比(按最终结论)')
        print('='*60)
        if ret_col in df.columns and '最终结论' in df.columns:
            df_ret = df.copy()
            df_ret[ret_col] = pd.to_numeric(df_ret[ret_col], errors='coerce')
            df_ret = df_ret.dropna(subset=[ret_col, '最终结论'])
            # 收益率分档(百分比)
            bins = [-999, -30, -20, -10, 0, 10, 20, 30, 999]
            labels = ['<-30%', '-30~-20%', '-20~-10%', '-10~0%', '0~10%', '10~20%', '20~30%', '>30%']
            df_ret['收益区间'] = pd.cut(df_ret[ret_col] * 100, bins=bins, labels=labels)
            # 交叉表: 收益区间 × 最终结论
            ret_crosstab = pd.crosstab(df_ret['收益区间'], df_ret['最终结论'], margins=True)
            # 占比(列内百分比)
            ret_pct = pd.crosstab(df_ret['收益区间'], df_ret['最终结论'], normalize='columns') * 100
            ret_pct = ret_pct.round(1)
            ret_pct['All'] = (pd.crosstab(df_ret['收益区间'], df_ret['最终结论'], normalize='all') * 100).round(1).sum(axis=1)
            print('数量分布:')
            print(ret_crosstab.to_string())
            print(f'\n占比分布(%):')
            print(ret_pct.to_string())
            # 写入Excel
            ret_crosstab.to_excel(writer, sheet_name='6b_收益区间占比', startrow=0)
            ret_pct.to_excel(writer, sheet_name='6b_收益区间占比', startrow=len(ret_crosstab) + 3)

        # ========== 7. 决策 vs 结论 交叉表 ==========
        print(f'\n{"="*60}')
        print('7. 定增决策 vs 最终结论 交叉表')
        print('='*60)
        if '定增决策' in df.columns and '最终结论' in df.columns:
            crosstab = pd.crosstab(df['定增决策'], df['最终结论'], margins=True)
            print(crosstab.to_string())
            crosstab.to_excel(writer, sheet_name='7_决策vs结论')

        # ========== 8. 通过的股票明细 ==========
        passed = df[df['最终结论'] == '通过']
        show_cols = [c for c in ['股票代码', '股票简称', '一级行业', '溢价率下限', '溢价率上限',
                                  '定增决策', ret_col, f'总分_{latest_year}' if score_cols else '总分',
                                  '成长能力_趋势'] if c in df.columns]
        if len(passed) > 0:
            print(f'\n{"="*60}')
            print(f'8. 通过的股票明细({len(passed)}只)')
            print('='*60)
            print(passed[show_cols].to_string(index=False))
            passed[show_cols].to_excel(writer, sheet_name='8_通过明细', index=False)

    print(f'\n✅ 透视分析已保存: {output_path}')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python analyze_screening_result.py <scored_excel_path>')
        sys.exit(1)
    analyze(sys.argv[1])
