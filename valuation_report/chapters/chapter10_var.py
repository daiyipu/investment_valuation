#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第十章 - VaR风险度量

功能：日收益率计算、多置信区间VaR/CVaR、最大回撤分析、VaR分布图、回撤分析图、风险-收益图谱
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from module_utils import add_title, add_paragraph, add_table_data, add_image, add_section_break


def _generate_return_volatility_scatter(windows_stats, save_path, stock_name='', font_prop=None):
    """生成风险-收益散点图"""
    fig, ax = plt.subplots(figsize=(10, 7))

    names = [w['name'] for w in windows_stats]
    returns = [w['annualized_return'] for w in windows_stats]
    volatilities = [w['annualized_volatility'] for w in windows_stats]

    scatter = ax.scatter(volatilities, returns, s=120, c=range(len(windows_stats)),
                         cmap='viridis', alpha=0.8, edgecolors='black', linewidth=0.5, zorder=5)

    for i, name in enumerate(names):
        ax.annotate(name, (volatilities[i], returns[i]),
                    textcoords="offset points", xytext=(8, 5),
                    fontsize=9, fontproperties=font_prop)

    ax.axhline(y=0, color='gray', linestyle='--', alpha=0.5)
    ax.set_xlabel('年化波动率(%)', fontproperties=font_prop, fontsize=12)
    ax.set_ylabel('年化收益率(%)', fontproperties=font_prop, fontsize=12)
    ax.set_title(f'{stock_name} 不同窗口期风险-收益图谱', fontproperties=font_prop, fontsize=14, fontweight='bold')
    ax.grid(True, alpha=0.3)
    for label in ax.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax.get_yticklabels():
        label.set_fontproperties(font_prop)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_chapter(context):
    """
    生成第十章：VaR风险度量

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

    price_data = context.get('price_data', pd.DataFrame())
    historical_valuation = context.get('historical_valuation', pd.DataFrame())
    daily_basic = context.get('daily_basic', {})

    current_price = float(daily_basic.get('close', 0))

    # ==================== 十、VaR风险度量 ====================
    add_title(document, '十、VaR风险度量', level=1)
    add_paragraph(document,
        '本章使用历史模拟法计算风险价值（VaR）和条件风险价值（CVaR），'
        '分析不同置信水平下的最大潜在损失，并进行最大回撤分析，'
        '全面评估投资标的的市场风险。')

    # ==================== 10.1 计算日收益率 ====================
    add_title(document, '10.1 日收益率分析', level=2)

    # 获取价格序列用于计算收益率
    prices = pd.Series(dtype=float)
    returns = pd.Series(dtype=float)

    if price_data is not None and not price_data.empty and 'close' in price_data.columns:
        price_data_sorted = price_data.sort_values('trade_date').reset_index(drop=True)
        prices = price_data_sorted['close'].astype(float)

        # 优先使用pct_chg列，否则从close计算
        if 'pct_chg' in price_data_sorted.columns:
            returns = price_data_sorted['pct_chg'].astype(float) / 100.0
        else:
            returns = prices.pct_change().dropna()
    elif historical_valuation is not None and not historical_valuation.empty and 'close' in historical_valuation.columns:
        hv_sorted = historical_valuation.sort_values('trade_date').reset_index(drop=True)
        prices = hv_sorted['close'].astype(float)
        returns = prices.pct_change().dropna()

    if returns.empty or len(returns) < 30:
        add_paragraph(document,
            '可用的价格数据不足（少于30个交易日），无法进行可靠的VaR分析。'
            '以下分析基于有限数据，结论仅供参考。')
        # 使用模拟数据作为后备
        if len(returns) > 0:
            add_paragraph(document, f'可用数据点数：{len(returns)}个交易日')
        else:
            add_paragraph(document, '无可用价格数据，跳过VaR分析。')
            context['results']['var_result'] = {
                'var_90': 0, 'var_95': 0, 'var_99': 0,
                'cvar_95': 0, 'cvar_99': 0, 'max_drawdown': 0,
            }
            add_section_break(document)
            return context

    add_paragraph(document, f'分析数据区间：共{len(returns)}个交易日')
    add_paragraph(document, f'日均收益率：{returns.mean():.4%}')
    add_paragraph(document, f'日收益率标准差：{returns.std():.4%}')
    add_paragraph(document, f'年化收益率：{returns.mean() * 252:.2%}')
    add_paragraph(document, f'年化波动率：{returns.std() * np.sqrt(252):.2%}')

    # 收益率统计表
    stats_headers = ['统计指标', '数值']
    stats_data = [
        ['日均收益率', f'{returns.mean():.4%}'],
        ['日均收益率标准差', f'{returns.std():.4%}'],
        ['最大日涨幅', f'{returns.max():.4%}'],
        ['最大日跌幅', f'{returns.min():.4%}'],
        ['偏度(Skewness)', f'{returns.skew():.4f}'],
        ['峰度(Kurtosis)', f'{returns.kurtosis():.4f}'],
        ['正收益天数', f'{(returns > 0).sum()}天({(returns > 0).mean():.1%})'],
        ['负收益天数', f'{(returns <= 0).sum()}天({(returns <= 0).mean():.1%})'],
    ]
    add_table_data(document, stats_headers, stats_data)

    # ==================== 10.2 多置信区间VaR/CVaR ====================
    add_paragraph(document, '')
    add_title(document, '10.2 多置信区间VaR/CVaR', level=2)
    add_paragraph(document,
        'VaR（Value at Risk）表示在给定置信水平下，投资可能遭受的最大损失。'
        'CVaR（Conditional VaR）又称期望短缺，表示在损失超过VaR时的平均损失，'
        '比VaR更保守，能更好地反映尾部风险。')
    add_paragraph(document,
        '本节采用历史模拟法（Historical Method），不假设收益率服从特定分布，'
        '直接从历史数据中提取分位数。')

    confidence_levels = [0.90, 0.95, 0.99]
    var_results = {}

    var_table_headers = ['置信水平', 'VaR(日)', 'VaR(年化)', 'CVaR(日)', 'CVaR(年化)', '含义']
    var_table_data = []

    for cl in confidence_levels:
        percentile = (1 - cl) * 100  # 对应的分位数
        var_value = np.percentile(returns, percentile)
        # CVaR = 所有低于VaR的收益率的均值
        tail_returns = returns[returns <= var_value]
        cvar_value = tail_returns.mean() if len(tail_returns) > 0 else var_value

        var_annual = var_value * np.sqrt(252)
        cvar_annual = cvar_value * np.sqrt(252)

        var_results[f'var_{int(cl * 100)}'] = var_value
        var_results[f'cvar_{int(cl * 100)}'] = cvar_value

        meaning_map = {
            0.90: '有10%的概率日损失超过此值',
            0.95: '有5%的概率日损失超过此值',
            0.99: '有1%的概率日损失超过此值',
        }

        var_table_data.append([
            f'{cl:.0%}',
            f'{var_value:.2%}',
            f'{var_annual:.2%}',
            f'{cvar_value:.2%}',
            f'{cvar_annual:.2%}',
            meaning_map[cl],
        ])

    add_table_data(document, var_table_headers, var_table_data)

    add_paragraph(document, '')
    add_paragraph(document, 'VaR/CVaR分析结论：', bold=True)

    var_95 = var_results.get('var_95', 0)
    cvar_95 = var_results.get('cvar_95', 0)
    var_99 = var_results.get('var_99', 0)
    cvar_99 = var_results.get('cvar_99', 0)

    add_paragraph(document,
        f'95% VaR为{var_95:.2%}（日），即有95%的把握日损失不超过{abs(var_95):.2%}。'
        f'对应年化VaR为{var_95 * np.sqrt(252):.2%}。')
    add_paragraph(document,
        f'95% CVaR为{cvar_95:.2%}（日），表示在最差5%的交易日中，'
        f'平均损失为{abs(cvar_95):.2%}，高于VaR的{abs(var_95):.2%}。')
    add_paragraph(document,
        f'CVaR与VaR的差异：95%水平下CVaR比VaR高{abs(cvar_95 - var_95):.2%}，'
        f'说明极端损失的尾部风险需要额外关注。')

    # ==================== 10.3 最大回撤分析 ====================
    add_paragraph(document, '')
    add_title(document, '10.3 最大回撤分析', level=2)
    add_paragraph(document,
        '最大回撤衡量从历史最高点到后续最低点的最大跌幅，反映投资可能面临的最严重亏损。')

    if len(prices) >= 2:
        peak = prices.cummax()
        drawdown = (prices - peak) / peak
        max_drawdown = drawdown.min()

        # 找到最大回撤的位置
        max_dd_end_idx = drawdown.idxmin()
        max_dd_end_price = prices.iloc[max_dd_end_idx]
        max_dd_peak_price = peak.iloc[max_dd_end_idx]

        # 找到回撤起始点（峰值位置）
        peak_before_dd = prices.iloc[:max_dd_end_idx + 1].idxmax()
        if isinstance(peak_before_dd, (int, np.integer)):
            max_dd_start_idx = int(peak_before_dd)
        else:
            max_dd_start_idx = prices.iloc[:max_dd_end_idx + 1].values.argmax()

        # 计算回撤持续期
        dd_duration = max_dd_end_idx - max_dd_start_idx

        # 计算当前回撤
        current_drawdown = drawdown.iloc[-1]

        add_paragraph(document, f'最大回撤幅度：{max_drawdown:.2%}')
        add_paragraph(document, f'最大回撤起始价：{max_dd_peak_price:.2f}元')
        add_paragraph(document, f'最大回撤终止价：{max_dd_end_price:.2f}元')
        add_paragraph(document, f'最大回撤持续期：{dd_duration}个交易日')
        add_paragraph(document, f'当前回撤水平：{current_drawdown:.2%}')

        dd_headers = ['回撤指标', '数值']
        dd_data = [
            ['最大回撤', f'{max_drawdown:.2%}'],
            ['最大回撤持续期', f'{dd_duration}个交易日'],
            ['当前回撤', f'{current_drawdown:.2%}'],
            ['回撤恢复所需涨幅', f'{-current_drawdown / (1 + current_drawdown):.2%}' if current_drawdown < 0 else '0.00%'],
        ]
        add_table_data(document, dd_headers, dd_data)

        # 回撤评级
        add_paragraph(document, '')
        add_paragraph(document, '回撤风险评估：', bold=True)
        if abs(max_drawdown) > 0.50:
            add_paragraph(document,
                f'历史最大回撤{max_drawdown:.2%}超过50%，风险极高。投资者需要有承受超过50%浮亏的心理准备。')
        elif abs(max_drawdown) > 0.30:
            add_paragraph(document,
                f'历史最大回撤{max_drawdown:.2%}在30%-50%之间，风险较高。建议控制仓位，做好风险管理。')
        elif abs(max_drawdown) > 0.15:
            add_paragraph(document,
                f'历史最大回撤{max_drawdown:.2%}在15%-30%之间，风险中等。合理的波动范围内。')
        else:
            add_paragraph(document,
                f'历史最大回撤{max_drawdown:.2%}低于15%，股价走势相对稳健。')
    else:
        max_drawdown = 0
        add_paragraph(document, '价格数据不足以进行回撤分析。')

    # ==================== 10.4 VaR分布图 ====================
    add_paragraph(document, '')
    add_title(document, '10.4 VaR分布图', level=2)

    var_chart_dict = {}
    for key in ['var_90', 'var_95', 'var_99', 'cvar_95', 'cvar_99']:
        if key in var_results:
            var_chart_dict[key] = var_results[key]

    var_chart_path = os.path.join(IMAGES_DIR, f'{stock_code.replace(".", "_")}_var_distribution.png')
    try:
        from module_utils import generate_var_distribution_chart
        generate_var_distribution_chart(returns, var_chart_dict, var_chart_path,
                                        title=f'{stock_name} VaR分析')
        add_image(document, var_chart_path)
    except Exception as e:
        add_paragraph(document, f'[VaR分布图生成失败: {e}]')

    add_paragraph(document, '')
    add_paragraph(document,
        '上图展示了日收益率的分布情况，彩色虚线标注了不同置信水平的VaR和CVaR位置。'
        'VaR线左侧的区域即为该置信水平下的尾部风险区域。')

    # ==================== 10.5 回撤分析图 ====================
    add_paragraph(document, '')
    add_title(document, '10.5 回撤分析图', level=2)

    dd_chart_path = os.path.join(IMAGES_DIR, f'{stock_code.replace(".", "_")}_drawdown.png')
    if len(prices) >= 2:
        try:
            from module_utils import generate_drawdown_chart
            generate_drawdown_chart(prices, dd_chart_path, title=f'{stock_name} 最大回撤分析')
            add_image(document, dd_chart_path)
        except Exception as e:
            add_paragraph(document, f'[回撤分析图生成失败: {e}]')
    else:
        add_paragraph(document, '价格数据不足，无法生成回撤分析图。')

    add_paragraph(document, '')
    add_paragraph(document,
        '上图上方为股价走势与历史最高价的对比，下方为各时点的回撤幅度。'
        '回撤幅度为负值区域反映了股价从峰值的下跌程度。')

    # ==================== 10.6 风险-收益图谱 ====================
    add_paragraph(document, '')
    add_title(document, '10.6 风险-收益图谱', level=2)
    add_paragraph(document,
        '不同窗口期的年化收益率与年化波动率构成风险-收益图谱，'
        '帮助投资者直观了解不同时间维度下的风险收益特征。')

    windows = [20, 60, 120, 250]
    windows_stats = []
    scatter_chart_path = os.path.join(IMAGES_DIR, f'{stock_code.replace(".", "_")}_risk_return.png')

    if len(returns) >= 20:
        for w in windows:
            if len(returns) >= w:
                rolling_ret = returns.rolling(w).sum()
                rolling_vol = returns.rolling(w).std() * np.sqrt(252)
                ann_ret = returns.tail(w).mean() * 252
                ann_vol = returns.tail(w).std() * np.sqrt(252)
                windows_stats.append({
                    'name': f'{w}日',
                    'annualized_return': ann_ret * 100,
                    'annualized_volatility': ann_vol * 100,
                })

        if windows_stats:
            # 风险-收益数据表
            rr_headers = ['时间窗口', '年化收益率(%)', '年化波动率(%)', '夏普比率(近似)']
            rr_data = []
            risk_free_annual = 0.0185  # 与config中的无风险利率一致
            for ws in windows_stats:
                sharpe = (ws['annualized_return'] / 100 - risk_free_annual) / (ws['annualized_volatility'] / 100) if ws['annualized_volatility'] > 0 else 0
                rr_data.append([
                    ws['name'],
                    f'{ws["annualized_return"]:.2f}',
                    f'{ws["annualized_volatility"]:.2f}',
                    f'{sharpe:.3f}'
                ])
            add_table_data(document, rr_headers, rr_data)

            add_paragraph(document, '')
            try:
                _generate_return_volatility_scatter(windows_stats, scatter_chart_path, stock_name, font_prop)
                add_image(document, scatter_chart_path)
            except Exception as e:
                add_paragraph(document, f'[风险-收益图谱生成失败: {e}]')

            add_paragraph(document, '')
            add_paragraph(document, '风险-收益分析结论：', bold=True)
            best_sharpe_ws = max(windows_stats,
                                 key=lambda x: (x['annualized_return'] / 100 - risk_free_annual) / (
                                     x['annualized_volatility'] / 100) if x['annualized_volatility'] > 0 else -999)
            add_paragraph(document,
                f'从不同窗口期来看，{best_sharpe_ws["name"]}窗口期的风险调整后收益（夏普比率）最优。')
    else:
        add_paragraph(document, '数据量不足，无法生成风险-收益图谱。')

    # ==================== 写入结果 ====================
    context['results']['var_result'] = {
        'var_90': var_results.get('var_90', 0),
        'var_95': var_results.get('var_95', 0),
        'var_99': var_results.get('var_99', 0),
        'cvar_95': var_results.get('cvar_95', 0),
        'cvar_99': var_results.get('cvar_99', 0),
        'max_drawdown': float(max_drawdown) if len(prices) >= 2 else 0,
    }

    add_section_break(document)
    return context
