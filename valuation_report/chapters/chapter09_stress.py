#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第九章 - 压力测试

功能：PE回归压力测试、极端经济情景压力测试、多因子极端测试、压力测试汇总
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from module_utils import add_title, add_paragraph, add_table_data, add_image, add_section_break


def _generate_stress_bar_chart(scenarios, save_path, stock_name='', font_prop=None):
    """生成压力测试柱状图"""
    fig, ax = plt.subplots(figsize=(12, 6))
    names = [s['name'] for s in scenarios]
    changes = [s['change_pct'] for s in scenarios]
    colors = ['#e74c3c' if c < 0 else '#2ecc71' for c in changes]

    y_pos = range(len(names))
    bars = ax.barh(y_pos, changes, color=colors, alpha=0.7, height=0.5)
    ax.axvline(x=0, color='black', linewidth=1.0)

    for bar, change in zip(bars, changes):
        ax.text(bar.get_width() + (1 if change >= 0 else -1),
                bar.get_y() + bar.get_height() / 2,
                f'{change:+.1f}%', ha='left' if change >= 0 else 'right',
                va='center', fontsize=9, fontproperties=font_prop)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(names, fontproperties=font_prop, fontsize=10)
    ax.set_xlabel('变动幅度(%)', fontproperties=font_prop, fontsize=12)
    ax.set_title(f'{stock_name}压力测试情景分析', fontproperties=font_prop, fontsize=14, fontweight='bold')
    ax.grid(True, alpha=0.3, axis='x')
    for label in ax.get_xticklabels():
        label.set_fontproperties(font_prop)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_chapter(context):
    """
    生成第九章：压力测试

    Parameters:
        context: dict 包含所有数据的上下文字典

    Returns:
        dict: 更新后的上下文字典
    """
    document = context['document']
    font_prop = context['font_prop']
    stock_name = context['stock_name']
    stock_code = context['stock_code']
    IMAGES_DIR = context['IMAGES_DIR']

    historical_valuation = context.get('historical_valuation', pd.DataFrame())
    daily_basic = context.get('daily_basic', {})
    wacc_result = context.get('wacc_result', {})
    price_data = context.get('price_data', pd.DataFrame())
    results = context.get('results', {})

    # 获取当前价格和估值数据
    current_price = float(daily_basic.get('close', 0))
    pe_ttm = float(daily_basic.get('pe_ttm', 0))
    pb = float(daily_basic.get('pb', 0))

    # 计算EPS
    eps = current_price / pe_ttm if pe_ttm > 0 else 0

    # 获取WACC
    wacc = float(wacc_result.get('wacc', 0.08)) if wacc_result else 0.08

    # 获取DCF结果
    dcf_result = results.get('dcf_result', {})
    dcf_summary = results.get('dcf_summary', {})
    # dcf_summary可能是list（chapter04输出）或dict
    if isinstance(dcf_summary, list) and dcf_summary:
        dcf_per_share = float(dcf_summary[0].get('per_share_value', 0)) if isinstance(dcf_summary[0], dict) else 0
    elif isinstance(dcf_summary, dict):
        dcf_per_share = float(dcf_summary.get('per_share_value', 0))
    else:
        dcf_per_share = float(results.get('dcf_avg_per_share', 0))

    # ==================== 九、压力测试 ====================
    add_title(document, '九、压力测试', level=1)
    add_paragraph(document,
        '本章节从多个维度模拟极端市场情况下的股价表现，包括PE估值回归测试、'
        '极端经济情景冲击、以及多重风险因子叠加的最差情景分析，'
        '全面评估投资标的在极端环境下的风险承受能力。')

    # 收集所有压力测试情景
    all_scenarios = []

    # ==================== 9.1 PE回归测试 ====================
    add_title(document, '9.1 PE回归测试', level=2)
    add_paragraph(document,
        'PE回归测试分析当PE估值回归到历史均值、中位数、四分位数等统计水平时，'
        '对应的目标价格。通过对比当前PE在历史分布中的位置，评估估值回调的风险。')

    # 计算历史PE统计量
    pe_stats = {}
    if historical_valuation is not None and not historical_valuation.empty and 'pe_ttm' in historical_valuation.columns:
        pe_series = historical_valuation['pe_ttm'].dropna()
        pe_series = pe_series[pe_series > 0]
        if len(pe_series) > 20:
            pe_stats = {
                'mean': float(pe_series.mean()),
                'median': float(pe_series.median()),
                'q1': float(pe_series.quantile(0.25)),
                'q3': float(pe_series.quantile(0.75)),
                'min': float(pe_series.min()),
                'max': float(pe_series.max()),
                'current': pe_ttm,
            }

    if pe_stats and eps > 0:
        add_paragraph(document, f'当前EPS（每股收益）：{eps:.4f}元，当前PE(TTM)：{pe_ttm:.2f}倍')

        # PE回归情景
        pe_scenarios = [
            ('PE回归均值', pe_stats['mean']),
            ('PE回归中位数', pe_stats['median']),
            ('PE回归Q1(25%)', pe_stats['q1']),
            ('PE回归Q3(75%)', pe_stats['q3']),
        ]

        pe_table_headers = ['情景', '目标PE(倍)', '目标价格(元)', '变动幅度(%)']
        pe_table_data = []

        for name, target_pe in pe_scenarios:
            target_price = eps * target_pe
            change_pct = (target_price - current_price) / current_price * 100 if current_price > 0 else 0
            pe_table_data.append([
                name,
                f'{target_pe:.2f}',
                f'{target_price:.2f}',
                f'{change_pct:+.2f}'
            ])
            all_scenarios.append({
                'name': name,
                'type': 'PE回归',
                'assumptions': f'PE从{pe_ttm:.2f}回归至{target_pe:.2f}',
                'target_price': target_price,
                'change_pct': change_pct,
            })

        # 添加当前基准行
        add_paragraph(document, '')
        add_table_data(document, pe_table_headers, pe_table_data)

        add_paragraph(document, '')
        add_paragraph(document, 'PE回归分析结论：', bold=True)

        # 判断当前PE在历史分布中的位置
        if pe_ttm > pe_stats['q3']:
            add_paragraph(document,
                f'当前PE({pe_ttm:.2f}倍)高于历史Q3({pe_stats["q3"]:.2f}倍)，估值处于历史较高水平，'
                f'存在较大的估值回调风险。若PE回归至历史均值({pe_stats["mean"]:.2f}倍)，'
                f'目标价为{eps * pe_stats["mean"]:.2f}元。')
        elif pe_ttm > pe_stats['median']:
            add_paragraph(document,
                f'当前PE({pe_ttm:.2f}倍)高于历史中位数({pe_stats["median"]:.2f}倍)但低于Q3，'
                f'估值处于中等偏高区间。若PE回归至中位数，目标价为{eps * pe_stats["median"]:.2f}元。')
        elif pe_ttm > pe_stats['q1']:
            add_paragraph(document,
                f'当前PE({pe_ttm:.2f}倍)处于历史Q1至中位数之间，估值处于合理区间，'
                f'PE回调风险相对有限。')
        else:
            add_paragraph(document,
                f'当前PE({pe_ttm:.2f}倍)低于历史Q1({pe_stats["q1"]:.2f}倍)，'
                f'估值处于历史较低水平，具有一定的安全边际。')
    else:
        add_paragraph(document, '历史PE数据不足，无法进行PE回归测试。使用替代方案：基于当前PE上下浮动。')
        if eps > 0 and pe_ttm > 0:
            for name, factor in [('PE上调20%', 1.2), ('PE下调20%', 0.8), ('PE下调40%', 0.6)]:
                target_pe = pe_ttm * factor
                target_price = eps * target_pe
                change_pct = (target_price - current_price) / current_price * 100 if current_price > 0 else 0
                all_scenarios.append({
                    'name': name,
                    'type': 'PE回归',
                    'assumptions': f'PE乘以{factor}',
                    'target_price': target_price,
                    'change_pct': change_pct,
                })

    # ==================== 9.2 极端经济情景 ====================
    add_paragraph(document, '')
    add_title(document, '9.2 极端经济情景', level=2)
    add_paragraph(document,
        '本节模拟三种极端宏观经济情景对估值的影响：通胀加剧、利率上行、经济衰退。'
        '每种情景分别影响折现率、WACC和盈利水平，进而影响每股价值。')

    # 定义极端经济情景
    economic_scenarios = [
        {
            'name': '通胀加剧(+3%)',
            'type': '经济情景',
            'assumptions': '通胀率上升3个百分点，折现率上升，PE压缩约15%',
            'wacc_adj': 0.03,      # WACC上升3%
            'pe_adj': 0.85,        # PE压缩15%
            'earnings_adj': 0.95,  # 盈利下降5%
        },
        {
            'name': '利率上行(+2%)',
            'type': '经济情景',
            'assumptions': '无风险利率上升2个百分点，WACC上升，估值压缩约10%',
            'wacc_adj': 0.02,
            'pe_adj': 0.90,
            'earnings_adj': 0.97,
        },
        {
            'name': '经济衰退(-20%)',
            'type': '经济情景',
            'assumptions': '经济衰退，企业盈利下降20%，PE压缩约25%',
            'wacc_adj': 0.01,
            'pe_adj': 0.75,
            'earnings_adj': 0.80,
        },
    ]

    eco_table_headers = ['情景', '情景假设', '调整后WACC', '调整后PE', '目标价格(元)', '变动幅度(%)']
    eco_table_data = []

    for scenario in economic_scenarios:
        adjusted_wacc = wacc + scenario['wacc_adj']
        adjusted_pe = pe_ttm * scenario['pe_adj']
        adjusted_eps = eps * scenario['earnings_adj']
        target_price = adjusted_eps * adjusted_pe
        change_pct = (target_price - current_price) / current_price * 100 if current_price > 0 else 0

        scenario['target_price'] = target_price
        scenario['change_pct'] = change_pct

        eco_table_data.append([
            scenario['name'],
            scenario['assumptions'],
            f'{adjusted_wacc:.2%}',
            f'{adjusted_pe:.2f}',
            f'{target_price:.2f}',
            f'{change_pct:+.2f}'
        ])
        all_scenarios.append({
            'name': scenario['name'],
            'type': scenario['type'],
            'assumptions': scenario['assumptions'],
            'target_price': target_price,
            'change_pct': change_pct,
        })

    add_table_data(document, eco_table_headers, eco_table_data)

    add_paragraph(document, '')
    add_paragraph(document, '极端经济情景分析：', bold=True)

    worst_eco = min(economic_scenarios, key=lambda x: x['change_pct'])
    add_paragraph(document,
        f'在三种极端经济情景中，{worst_eco["name"]}情景对股价冲击最大，'
        f'目标价格{worst_eco["target_price"]:.2f}元，变动幅度{worst_eco["change_pct"]:+.2f}%。')

    # 计算各情景下DCF每股价值的影响
    if dcf_per_share > 0:
        add_paragraph(document, '')
        add_paragraph(document, '极端情景对DCF估值的影响：', bold=True)
        dcf_stress_headers = ['情景', '基准DCF(元/股)', '调整后DCF(元/股)', '变动幅度(%)']
        dcf_stress_data = []
        for scenario in economic_scenarios:
            # DCF受WACC和增长率影响，简化计算
            wacc_ratio = wacc / (wacc + scenario['wacc_adj']) if (wacc + scenario['wacc_adj']) > 0 else 1
            adjusted_dcf = dcf_per_share * wacc_ratio * scenario['earnings_adj']
            change = (adjusted_dcf - dcf_per_share) / dcf_per_share * 100
            dcf_stress_data.append([
                scenario['name'],
                f'{dcf_per_share:.2f}',
                f'{adjusted_dcf:.2f}',
                f'{change:+.2f}'
            ])
        add_table_data(document, dcf_stress_headers, dcf_stress_data)

    # ==================== 9.3 多因子极端测试 ====================
    add_paragraph(document, '')
    add_title(document, '9.3 多因子极端测试', level=2)
    add_paragraph(document,
        '多因子极端测试将最不利的因子组合叠加：最差PE水平 + 最差盈利增长 + 最差WACC，'
        '模拟"完美风暴"情景下标的的估值表现。')

    # 确定极端参数
    worst_pe = pe_stats.get('q1', pe_ttm * 0.6) if pe_stats else pe_ttm * 0.6
    worst_earnings_factor = 0.75   # 盈利下降25%
    worst_wacc_addition = 0.05     # WACC额外上升5%

    worst_eps = eps * worst_earnings_factor
    worst_target_pe = worst_pe
    worst_target_price = worst_eps * worst_target_pe
    worst_change_pct = (worst_target_price - current_price) / current_price * 100 if current_price > 0 else 0
    worst_adjusted_wacc = wacc + worst_wacc_addition

    add_paragraph(document, '多因子极端参数：', bold=True)
    extreme_params_headers = ['参数', '当前值', '极端值', '变动']
    extreme_params_data = [
        ['PE(TTM)', f'{pe_ttm:.2f}倍', f'{worst_target_pe:.2f}倍',
         f'{(worst_target_pe / pe_ttm - 1) * 100:+.1f}%' if pe_ttm > 0 else 'N/A'],
        ['EPS', f'{eps:.4f}元', f'{worst_eps:.4f}元', f'{(worst_earnings_factor - 1) * 100:+.1f}%'],
        ['WACC', f'{wacc:.2%}', f'{worst_adjusted_wacc:.2%}', f'+{worst_wacc_addition:.2%}'],
    ]
    add_table_data(document, extreme_params_headers, extreme_params_data)

    add_paragraph(document, '')
    add_paragraph(document, '多因子极端测试结果：', bold=True)
    multi_factor_headers = ['指标', '当前', '极端情景', '变动']
    multi_factor_data = [
        ['每股收益(EPS)', f'{eps:.4f}元', f'{worst_eps:.4f}元',
         f'{(worst_earnings_factor - 1) * 100:+.1f}%'],
        ['PE估值', f'{pe_ttm:.2f}倍', f'{worst_target_pe:.2f}倍',
         f'{(worst_target_pe / pe_ttm - 1) * 100:+.1f}%' if pe_ttm > 0 else 'N/A'],
        ['目标价格', f'{current_price:.2f}元', f'{worst_target_price:.2f}元',
         f'{worst_change_pct:+.2f}%'],
    ]
    add_table_data(document, multi_factor_headers, multi_factor_data)

    # DCF极端
    if dcf_per_share > 0:
        worst_dcf = dcf_per_share * (wacc / worst_adjusted_wacc) * worst_earnings_factor
        worst_dcf_change = (worst_dcf - dcf_per_share) / dcf_per_share * 100
        add_paragraph(document,
            f'极端情景下DCF每股价值：{worst_dcf:.2f}元（基准{dcf_per_share:.2f}元，变动{worst_dcf_change:+.2f}%）')

    all_scenarios.append({
        'name': '多因子极端(最差PE+最差增长+最差WACC)',
        'type': '多因子极端',
        'assumptions': f'PE={worst_target_pe:.2f},EPS降幅{(1 - worst_earnings_factor) * 100:.0f}%,WACC+{worst_wacc_addition:.0%}',
        'target_price': worst_target_price,
        'change_pct': worst_change_pct,
    })

    # ==================== 9.4 压力测试汇总表 ====================
    add_paragraph(document, '')
    add_title(document, '9.4 压力测试汇总', level=2)
    add_paragraph(document, '下表汇总所有压力测试情景的结果：')

    summary_headers = ['情景名称', '类型', '情景假设', '目标价格(元)', '变动幅度(%)']
    summary_data = []
    for s in all_scenarios:
        summary_data.append([
            s['name'],
            s['type'],
            s['assumptions'],
            f'{s["target_price"]:.2f}',
            f'{s["change_pct"]:+.2f}'
        ])
    add_table_data(document, summary_headers, summary_data)

    # 生成压力测试汇总图
    add_paragraph(document, '')
    chart_path = os.path.join(IMAGES_DIR, f'{stock_code.replace(".", "_")}_stress_test.png')
    try:
        _generate_stress_bar_chart(all_scenarios, chart_path, stock_name, font_prop)
        add_image(document, chart_path)
    except Exception as e:
        add_paragraph(document, f'[压力测试图表生成失败: {e}]')

    # 压力测试结论
    add_paragraph(document, '')
    add_paragraph(document, '压力测试综合结论：', bold=True)

    negative_scenarios = [s for s in all_scenarios if s['change_pct'] < 0]
    worst_scenario = min(all_scenarios, key=lambda x: x['change_pct']) if all_scenarios else None

    if worst_scenario:
        add_paragraph(document,
            f'在所有{len(all_scenarios)}个压力测试情景中，{len(negative_scenarios)}个情景出现负收益。'
            f'最严重情景为"{worst_scenario["name"]}"，目标价格{worst_scenario["target_price"]:.2f}元，'
            f'变动幅度{worst_scenario["change_pct"]:+.2f}%。')

    if len(negative_scenarios) > len(all_scenarios) * 0.5:
        add_paragraph(document,
            '超过半数压力测试情景出现负收益，表明当前估值在极端情况下存在较大回调风险，'
            '建议投资者充分评估下行风险。')
    elif len(negative_scenarios) > 0:
        add_paragraph(document,
            '部分压力测试情景出现负收益，但整体风险可控。投资者应关注极端经济情景下的潜在损失。')
    else:
        add_paragraph(document,
            '所有压力测试情景均未出现负收益，表明当前估值具有较强的安全边际。')

    # 将结果写入context
    context['results']['stress_scenarios'] = all_scenarios

    add_section_break(document)
    return context
