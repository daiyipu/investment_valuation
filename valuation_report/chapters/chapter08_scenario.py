#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第八章 - 情景分析

包含:
- 8.1 市场PE情景: 牛市/中性/熊市
- 8.2 行业PE情景: 高/中/低
- 8.3 个股PE情景: 乐观/中性/悲观
- 8.4 3x3x5矩阵: market(3) x industry(3) x individual(5 levels)
- 8.5 情景柱状图
- 8.6 历史情景回溯
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

from module_utils import (
    add_title, add_paragraph, add_table_data, add_image, add_section_break,
    generate_scenario_bar_chart,
)


def _get_eps(context):
    """从daily_basic或financial_statements推算每股收益。"""
    daily_basic = context.get('daily_basic', {})
    pe = daily_basic.get('pe_ttm', None)
    close = daily_basic.get('close', None)
    if pe and close and pe > 0:
        return close / pe

    fs = context.get('financial_statements', {})
    income = fs.get('income', pd.DataFrame())
    if not income.empty and 'total_share' in daily_basic:
        total_share = daily_basic.get('total_share', None)
        if total_share and total_share > 0:
            net_col = 'netprofit_incl_min_int_inc' if 'netprofit_incl_min_int_inc' in income.columns else 'net_profit'
            if net_col in income.columns:
                latest_net = income.iloc[0][net_col]
                return float(latest_net) / float(total_share)
    return None


def _find_closest_pe_periods(historical_valuation, current_pe, top_n=5):
    """在历史估值中找到PE最接近当前PE的时期。"""
    if historical_valuation is None or historical_valuation.empty:
        return []

    hv = historical_valuation.copy()
    hv = hv[hv['pe_ttm'] > 0].dropna(subset=['pe_ttm'])

    if hv.empty or current_pe is None or current_pe <= 0:
        return []

    hv['pe_diff'] = np.abs(hv['pe_ttm'] - current_pe)
    hv_sorted = hv.sort_values('pe_diff').head(top_n)

    results = []
    for _, row in hv_sorted.iterrows():
        results.append({
            'trade_date': str(row.get('trade_date', '')),
            'pe_ttm': float(row.get('pe_ttm', 0)),
            'pb': float(row.get('pb', 0)) if pd.notna(row.get('pb')) else None,
            'close': float(row.get('close', 0)) if pd.notna(row.get('close')) else None,
            'pe_diff': float(row.get('pe_diff', 0)),
        })
    return results


def generate_chapter(context):
    """生成第八章：情景分析

    Args:
        context: dict with keys as specified in project main.py

    Returns:
        context (updated, results['scenarios'] populated)
    """
    document = context['document']
    font_prop = context['font_prop']
    pro = context.get('pro')
    config = context.get('config', {})
    daily_basic = context.get('daily_basic', {})
    historical_valuation = context.get('historical_valuation')
    sw_industry = context.get('sw_industry', {})
    peer_valuation = context.get('peer_valuation')
    IMAGES_DIR = context['IMAGES_DIR']
    results = context.setdefault('results', {})

    stock_name = context['stock_name']
    stock_code = context['stock_code']

    current_pe = daily_basic.get('pe_ttm', None)
    current_price = daily_basic.get('close', None)
    eps = _get_eps(context)

    # ==================== 八、情景分析 ====================
    add_title(document, '八、情景分析', level=1)
    add_paragraph(document, (
        '本章节通过构建不同PE情景，分析市场、行业、个股三个维度对估值的影响，'
        '形成完整的3x3x5（共45个）情景矩阵。'
    ))

    if eps and eps > 0:
        add_paragraph(document, f'当前PE(TTM): {current_pe:.2f}倍, 每股收益(EPS): {eps:.4f}元, 当前价格: {current_price:.2f}元')
    else:
        add_paragraph(document, f'当前PE(TTM): {current_pe}, 当前价格: {current_price}')
        add_paragraph(document, 'EPS数据不完整，部分情景分析可能受限。')

    # ------------------------------------------------------------------
    # 8.1  市场PE情景: 牛市/中性/熊市
    # ------------------------------------------------------------------
    add_title(document, '8.1 市场PE情景分析', level=2)
    add_paragraph(document, (
        '根据市场整体环境，设定三种市场PE情景：'
        '牛市（PE x 1.2）、中性（PE x 1.0）、熊市（PE x 0.8）。'
    ))

    market_factors = {'牛市': 1.2, '中性': 1.0, '熊市': 0.8}
    market_scenarios = []

    if current_pe and current_pe > 0 and eps and eps > 0:
        headers_mkt = ['市场情景', 'PE因子', '情景PE', '目标价(元)', '涨跌幅']
        data_mkt = []
        for name, factor in market_factors.items():
            scenario_pe = current_pe * factor
            target_price = scenario_pe * eps
            change_pct = (target_price / current_price - 1) * 100 if current_price and current_price > 0 else 0
            data_mkt.append([
                name,
                f'{factor:.1f}x',
                f'{scenario_pe:.2f}',
                f'{target_price:.2f}',
                f'{change_pct:+.1f}%',
            ])
            market_scenarios.append({
                'name': f'市场-{name}',
                'type': 'bull' if factor > 1 else 'base' if factor == 1 else 'bear',
                'factor': factor,
                'pe': scenario_pe,
                'value': target_price,
                'change_pct': change_pct,
            })
        add_table_data(document, headers_mkt, data_mkt)
    else:
        add_paragraph(document, 'PE或EPS数据不足，跳过市场情景分析。')

    # ------------------------------------------------------------------
    # 8.2  行业PE情景: 高/中/低
    # ------------------------------------------------------------------
    add_title(document, '8.2 行业PE情景分析', level=2)
    add_paragraph(document, (
        '基于申万三级行业指数PE，设定高/中/低三种行业估值情景。'
    ))

    industry_pe = None
    industry_scenarios = []

    # 尝试获取行业PE
    l3_code = sw_industry.get('l3_code', '')
    l3_name = sw_industry.get('l3_name', '')

    if l3_code and pro is not None:
        try:
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=30)).strftime('%Y%m%d')
            df_sw = pro.sw_daily(ts_code=l3_code, start_date=start_date, end_date=end_date)
            if df_sw is not None and not df_sw.empty:
                industry_pe = float(df_sw.iloc[-1].get('pe', 0))
        except Exception as e:
            print(f"  获取行业PE失败: {e}")

    # 备选：从peer_valuation计算
    if (industry_pe is None or industry_pe == 0) and peer_valuation is not None and not peer_valuation.empty:
        pv_pe = peer_valuation[(peer_valuation['pe_ttm'] > 0) & (peer_valuation['pe_ttm'] < 500)]
        if not pv_pe.empty:
            industry_pe = float(pv_pe['pe_ttm'].median())

    if industry_pe and industry_pe > 0:
        add_paragraph(document, f'行业PE(中位数): {industry_pe:.2f}倍 (来源: {l3_name if l3_name else "同行中位数"})')

        industry_factors = {'高': 1.2, '中': 1.0, '低': 0.8}
        headers_ind = ['行业情景', 'PE因子', '行业PE', '个股目标PE', '目标价(元)']
        data_ind = []

        for name, factor in industry_factors.items():
            scenario_ind_pe = industry_pe * factor
            # 个股目标价 = 行业情景PE x EPS（假设个股回归行业估值）
            if eps and eps > 0:
                target_price = scenario_ind_pe * eps
                data_ind.append([
                    name,
                    f'{factor:.1f}x',
                    f'{scenario_ind_pe:.2f}',
                    f'{scenario_ind_pe:.2f}',
                    f'{target_price:.2f}',
                ])
                industry_scenarios.append({
                    'name': f'行业-{name}',
                    'type': 'bull' if factor > 1 else 'base' if factor == 1 else 'bear',
                    'factor': factor,
                    'pe': scenario_ind_pe,
                    'value': target_price,
                })
            else:
                data_ind.append([name, f'{factor:.1f}x', f'{scenario_ind_pe:.2f}', '-', '-'])

        add_table_data(document, headers_ind, data_ind)
    else:
        add_paragraph(document, '行业PE数据不可用，跳过行业情景分析。')

    # ------------------------------------------------------------------
    # 8.3  个股PE情景: 乐观/中性/悲观
    # ------------------------------------------------------------------
    add_title(document, '8.3 个股PE情景分析', level=2)
    add_paragraph(document, (
        '基于个股历史PE分布，设定乐观/中性/悲观三种个股估值情景。'
    ))

    stock_scenarios = []

    if historical_valuation is not None and not historical_valuation.empty:
        hv_pe = historical_valuation['pe_ttm'].dropna()
        hv_pe = hv_pe[(hv_pe > 0) & (hv_pe < 500)]

        if not hv_pe.empty and eps and eps > 0:
            p10 = float(hv_pe.quantile(0.10))
            p25 = float(hv_pe.quantile(0.25))
            p50 = float(hv_pe.quantile(0.50))
            p75 = float(hv_pe.quantile(0.75))
            p90 = float(hv_pe.quantile(0.90))

            stock_levels = {
                '悲观(25%分位)': p25,
                '中性(50%分位)': p50,
                '乐观(75%分位)': p75,
            }

            headers_stock = ['个股情景', '历史PE分位', '目标PE', '目标价(元)', '涨跌幅']
            data_stock = []
            for name, scenario_pe in stock_levels.items():
                target_price = scenario_pe * eps
                change_pct = (target_price / current_price - 1) * 100 if current_price and current_price > 0 else 0
                data_stock.append([
                    name,
                    f'{scenario_pe:.2f}',
                    f'{scenario_pe:.2f}',
                    f'{target_price:.2f}',
                    f'{change_pct:+.1f}%',
                ])
                stock_scenarios.append({
                    'name': f'个股-{name}',
                    'type': 'bull' if scenario_pe > (current_pe or 0) else 'base' if abs(scenario_pe - (current_pe or 0)) / max((current_pe or 1), 1) < 0.1 else 'bear',
                    'pe': scenario_pe,
                    'value': target_price,
                    'change_pct': change_pct,
                })
            add_table_data(document, headers_stock, data_stock)

            # 显示历史PE分布
            add_paragraph(document, '')
            add_paragraph(document, f'历史PE分布: 10%={p10:.2f}, 25%={p25:.2f}, 50%={p50:.2f}, 75%={p75:.2f}, 90%={p90:.2f}')
        else:
            add_paragraph(document, '历史PE数据不足，跳过个股情景分析。')
    else:
        add_paragraph(document, '历史估值数据不足，跳过个股情景分析。')

    # ------------------------------------------------------------------
    # 8.4  3x3x5矩阵
    # ------------------------------------------------------------------
    add_title(document, '8.4 综合情景矩阵 (3x3x5)', level=2)
    add_paragraph(document, (
        '构建市场(3) x 行业(3) x 个股(5) = 45个情景的完整矩阵。'
        '市场: 牛市/中性/熊市 (PE x 1.2/1.0/0.8)；'
        '行业: 高/中/低 (行业PE x 1.2/1.0/0.8)；'
        f'个股: 基于历史PE分位数的5个水平。'
    ))

    all_scenarios = []

    if eps and eps > 0 and current_pe and current_pe > 0:
        # 市场维度
        mkt_levels = [('牛市', 1.2), ('中性', 1.0), ('熊市', 0.8)]

        # 行业维度
        ind_pe_base = industry_pe if industry_pe and industry_pe > 0 else current_pe
        ind_levels = [('高', 1.2), ('中', 1.0), ('低', 0.8)]

        # 个股维度（5个历史分位水平）
        individual_levels = []
        if historical_valuation is not None and not historical_valuation.empty:
            hv_pe = historical_valuation['pe_ttm'].dropna()
            hv_pe = hv_pe[(hv_pe > 0) & (hv_pe < 500)]
            if not hv_pe.empty:
                for q, label in [(0.10, '极度悲观'), (0.25, '悲观'), (0.50, '中性'), (0.75, '乐观'), (0.90, '极度乐观')]:
                    individual_levels.append((label, float(hv_pe.quantile(q))))

        if not individual_levels:
            # 回退到基于当前PE的5个水平
            for factor, label in [(0.7, '极度悲观'), (0.85, '悲观'), (1.0, '中性'), (1.15, '乐观'), (1.3, '极度乐观')]:
                individual_levels.append((label, current_pe * factor))

        # 生成情景矩阵
        scenario_headers = ['编号', '市场情景', '行业情景', '个股情景', '综合PE', '目标价(元)', '涨跌幅']
        scenario_data = []

        idx = 1
        for mkt_name, mkt_factor in mkt_levels:
            for ind_name, ind_factor in ind_levels:
                for ind_label, ind_pe in individual_levels:
                    # 综合PE = 个股历史PE分位数 x 市场因子 x 行业因子微调
                    # 市场因子影响整体估值水平
                    # 行业因子通过行业PE与个股PE的比值调整
                    if ind_pe_base > 0:
                        ind_ratio = (ind_pe_base * ind_factor) / ind_pe_base
                    else:
                        ind_ratio = ind_factor

                    composite_pe = ind_pe * mkt_factor * (0.5 + 0.5 * ind_ratio)
                    target_price = composite_pe * eps
                    change_pct = (target_price / current_price - 1) * 100 if current_price and current_price > 0 else 0

                    scenario_data.append([
                        f'{idx:02d}',
                        mkt_name,
                        ind_name,
                        ind_label,
                        f'{composite_pe:.2f}',
                        f'{target_price:.2f}',
                        f'{change_pct:+.1f}%',
                    ])

                    all_scenarios.append({
                        'id': idx,
                        'name': f'{mkt_name}-{ind_name}-{ind_label}',
                        'market': mkt_name,
                        'industry': ind_name,
                        'individual': ind_label,
                        'composite_pe': float(composite_pe),
                        'value': float(target_price),
                        'change_pct': float(change_pct),
                        'type': 'bull' if change_pct > 10 else 'base' if abs(change_pct) <= 10 else 'bear',
                    })
                    idx += 1

        if scenario_data:
            add_table_data(document, scenario_headers, scenario_data)
            add_paragraph(document, '')
            add_paragraph(document, f'共生成 {len(all_scenarios)} 个情景。')
            add_paragraph(document, f'最高目标价: {max(s["value"] for s in all_scenarios):.2f}元 ({max(s["name"] for s in all_scenarios if s["value"] == max(ss["value"] for ss in all_scenarios))})')
            add_paragraph(document, f'最低目标价: {min(s["value"] for s in all_scenarios):.2f}元 ({min(s["name"] for s in all_scenarios if s["value"] == min(ss["value"] for ss in all_scenarios))})')
            add_paragraph(document, f'中位数目标价: {np.median([s["value"] for s in all_scenarios]):.2f}元')
        else:
            add_paragraph(document, '无法生成情景矩阵（数据不完整）。')
    else:
        add_paragraph(document, 'PE或EPS数据不足，无法生成综合情景矩阵。')

    # ------------------------------------------------------------------
    # 8.5  情景柱状图
    # ------------------------------------------------------------------
    add_title(document, '8.5 情景分析柱状图', level=2)

    # 汇总9个主要情景（3市场 x 3行业，取个股中性水平）
    summary_scenarios = []
    if all_scenarios:
        # 取个股中性的情景
        neutral_individual = [s for s in all_scenarios if s['individual'] == '中性']
        if neutral_individual:
            summary_scenarios = neutral_individual
        else:
            # 取前9个
            summary_scenarios = all_scenarios[:9]

    chart_path = os.path.join(IMAGES_DIR, '08_scenario_bar_chart.png')
    if summary_scenarios:
        try:
            chart_scenarios = [
                {'name': s['name'], 'value': s['value'], 'type': s['type']}
                for s in summary_scenarios
            ]
            generate_scenario_bar_chart(
                chart_scenarios, chart_path,
                title=f'{stock_name} 情景分析目标价',
            )
            if os.path.exists(chart_path):
                add_paragraph(document, '图表8.1 情景分析目标价（个股中性水平）')
                add_image(document, chart_path)
        except Exception as e:
            print(f"  生成情景柱状图失败: {e}")

        # 额外绘制全部情景分布图
        if len(all_scenarios) > 9:
            try:
                fig, ax = plt.subplots(figsize=(14, 8))
                values = [s['value'] for s in all_scenarios]
                colors = ['#27ae60' if s['change_pct'] > 10 else '#f39c12' if abs(s['change_pct']) <= 10 else '#e74c3c'
                          for s in all_scenarios]
                sorted_indices = np.argsort(values)
                sorted_names = [all_scenarios[i]['name'] for i in sorted_indices]
                sorted_values = [all_scenarios[i]['value'] for i in sorted_indices]
                sorted_colors = [colors[i] for i in sorted_indices]

                ax.barh(range(len(sorted_names)), sorted_values, color=sorted_colors, alpha=0.7, height=0.7)
                if current_price:
                    ax.axvline(x=current_price, color='black', linestyle='--', linewidth=2,
                               label=f'当前价 {current_price:.2f}')

                ax.set_yticks(range(len(sorted_names)))
                ax.set_yticklabels(sorted_names, fontproperties=font_prop, fontsize=7)
                ax.set_xlabel('目标价(元)', fontproperties=font_prop, fontsize=12)
                ax.set_title(f'{stock_name} 全部情景目标价分布 (n={len(all_scenarios)})',
                             fontproperties=font_prop, fontsize=14, fontweight='bold')
                ax.legend(prop=font_prop)
                ax.grid(True, alpha=0.3, axis='x')
                for label in ax.get_xticklabels():
                    label.set_fontproperties(font_prop)
                plt.tight_layout()
                full_chart = os.path.join(IMAGES_DIR, '08_scenario_full_distribution.png')
                plt.savefig(full_chart, dpi=150, bbox_inches='tight')
                plt.close()
                if os.path.exists(full_chart):
                    add_paragraph(document, '图表8.2 全部45情景目标价分布')
                    add_image(document, full_chart)
            except Exception as e:
                print(f"  生成全部情景分布图失败: {e}")
    else:
        add_paragraph(document, '情景数据不足，无法生成柱状图。')

    # ------------------------------------------------------------------
    # 8.6  历史情景回溯
    # ------------------------------------------------------------------
    add_title(document, '8.6 历史情景回溯', level=2)
    add_paragraph(document, (
        '在历史估值数据中，找到PE最接近当前PE的5个时期，分析后续走势作为参考。'
    ))

    if historical_valuation is not None and not historical_valuation.empty and current_pe:
        closest_periods = _find_closest_pe_periods(historical_valuation, current_pe, top_n=5)

        if closest_periods:
            headers_hist = ['交易日期', '当时PE', '当时PB', '当时收盘价', '与当前PE差异']
            data_hist = []
            for p in closest_periods:
                data_hist.append([
                    p['trade_date'],
                    f'{p["pe_ttm"]:.2f}',
                    f'{p["pb"]:.2f}' if p.get('pb') else '-',
                    f'{p["close"]:.2f}' if p.get('close') else '-',
                    f'{p["pe_diff"]:.2f}',
                ])
            add_table_data(document, headers_hist, data_hist)

            add_paragraph(document, '')
            add_paragraph(document, '历史相似PE时期分析:')

            hv = historical_valuation.copy()
            hv = hv[hv['pe_ttm'] > 0].dropna(subset=['pe_ttm'])
            hv = hv.sort_values('trade_date').reset_index(drop=True)

            for p in closest_periods:
                trade_date = p['trade_date']
                idx_match = hv[hv['trade_date'] == trade_date].index
                if len(idx_match) > 0:
                    idx_pos = idx_match[0]
                    pe_at_time = p['pe_ttm']
                    close_at_time = p.get('close', None)

                    # 后续60/120日走势
                    add_paragraph(document, f'- {trade_date}: PE={pe_at_time:.2f}')
                    for forward_days, label in [(20, '20日'), (60, '60日'), (120, '120日')]:
                        future_idx = idx_pos + forward_days
                        if future_idx < len(hv):
                            future_close = hv.iloc[future_idx].get('close', None)
                            future_pe = hv.iloc[future_idx].get('pe_ttm', None)
                            if close_at_time and close_at_time > 0 and future_close:
                                ret = (future_close / close_at_time - 1) * 100
                                add_paragraph(document, (
                                    f'  后续{label}: 收盘价={future_close:.2f}, '
                                    f'PE={future_pe:.2f}, '
                                    f'收益率={ret:+.1f}%'
                                ))
                        else:
                            add_paragraph(document, f'  后续{label}: 数据不可用')
        else:
            add_paragraph(document, '未找到与当前PE相近的历史时期。')
    else:
        add_paragraph(document, '历史估值数据不足，跳过历史情景回溯。')

    add_section_break(document)

    # ==================== 保存结果 ====================
    context['results']['scenarios'] = all_scenarios

    return context
