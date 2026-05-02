#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第六章 - 敏感性分析

包含:
- 6.1 WACC x 增长率 敏感性矩阵 (9x9)
- 6.2 热力图
- 6.3 时间窗口分析: 20/60/120/250日历史PE分位数
- 6.4 折价敏感性: 不同折价率(0.7-1.0)下的目标价
- 6.5 波动率敏感性: 不同波动率对估值的影响
- 6.6 龙卷风图
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from module_utils import (
    add_title, add_paragraph, add_table_data, add_image, add_section_break,
    generate_heatmap, generate_tornado_chart,
)


def _get_base_fcf(context):
    """从 context 中提取基础 FCF（自由现金流，单位：万元）。

    优先从 results['dcf_result'] 获取；若不存在，则从 financial_statements
    近3年经营现金流-资本开支取均值。
    """
    results = context.get('results', {})
    dcf = results.get('dcf_result', {})
    if dcf and dcf.get('last_fcf') is not None:
        return float(dcf['last_fcf'])

    fs = context.get('financial_statements', {})
    cashflow = fs.get('cashflow', pd.DataFrame())
    if cashflow.empty:
        return None

    try:
        # 经营活动现金流量净额 - 资本支出，Tushare单位为元，转为万元
        ocf_col = 'n_cashflow_act' if 'n_cashflow_act' in cashflow.columns else None
        capex_col = 'c_pay_acq_const_fiolta' if 'c_pay_acq_const_fiolta' in cashflow.columns else None
        if ocf_col and capex_col:
            ocf = cashflow[ocf_col].dropna()
            capex = cashflow[capex_col].abs().dropna()
            min_len = min(len(ocf), len(capex), 3)
            if min_len > 0:
                fcf_vals = (ocf.iloc[:min_len].values - capex.iloc[:min_len].values) / 10000
                return float(np.mean(fcf_vals))
    except Exception:
        pass
    return None


def _compute_dcf_per_share(fcf, wacc, growth, terminal_growth, n_years, total_shares):
    """简化DCF模型：Gordon Growth Model 终值。"""
    if fcf is None or wacc is None or total_shares is None or total_shares == 0:
        return None
    if wacc <= growth:
        return None
    pv_fcf = 0.0
    for yr in range(1, n_years + 1):
        cf = fcf * (1 + growth) ** yr
        pv_fcf += cf / (1 + wacc) ** yr
    terminal_value = fcf * (1 + growth) ** n_years * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_terminal = terminal_value / (1 + wacc) ** n_years
    total_value = pv_fcf + pv_terminal
    return total_value / total_shares


def generate_chapter(context):
    """生成第六章：敏感性分析

    Args:
        context: dict with keys as specified in project main.py

    Returns:
        context (updated, results['sensitivity_matrix'] populated)
    """
    document = context['document']
    font_prop = context['font_prop']
    config = context.get('config', {})
    daily_basic = context.get('daily_basic', {})
    historical_valuation = context.get('historical_valuation')
    price_data = context.get('price_data')
    wacc_result = context.get('wacc_result', {})
    IMAGES_DIR = context['IMAGES_DIR']
    results = context.setdefault('results', {})

    # ==================== 六、敏感性分析 ====================
    add_title(document, '六、敏感性分析', level=1)
    add_paragraph(document, (
        '本章节通过调整关键参数（WACC、增长率、折价率、波动率等），'
        '分析各参数变动对估值结果的影响程度，识别核心敏感因子。'
    ))

    # ------------------------------------------------------------------
    # 6.1  WACC x 增长率 敏感性矩阵
    # ------------------------------------------------------------------
    add_title(document, '6.1 WACC x 增长率 敏感性矩阵', level=2)

    sens_cfg = config.get('sensitivity', {})
    dcf_cfg = config.get('dcf', {})

    wacc_min = sens_cfg.get('wacc_range', [0.06, 0.15])[0]
    wacc_max = sens_cfg.get('wacc_range', [0.06, 0.15])[1]
    wacc_steps = sens_cfg.get('wacc_steps', 9)
    growth_min = sens_cfg.get('growth_range', [-0.05, 0.20])[0]
    growth_max = sens_cfg.get('growth_range', [-0.05, 0.20])[1]
    growth_steps = sens_cfg.get('growth_steps', 9)
    terminal_growth = dcf_cfg.get('terminal_growth', 0.03)
    n_years = context.get('results', {}).get('dcf_result', {}).get('forecast_years', dcf_cfg.get('forecast_years', 10))

    base_fcf = _get_base_fcf(context)
    total_shares = daily_basic.get('total_share', None)

    wacc_values = np.linspace(wacc_min, wacc_max, wacc_steps)
    growth_values = np.linspace(growth_min, growth_max, growth_steps)

    matrix = np.full((wacc_steps, growth_steps), np.nan)
    wacc_labels = [f'{w * 100:.2f}%' for w in wacc_values]
    growth_labels = [f'{g * 100:.2f}%' for g in growth_values]

    if base_fcf is not None and total_shares is not None and total_shares > 0:
        for i, w in enumerate(wacc_values):
            for j, g in enumerate(growth_values):
                val = _compute_dcf_per_share(
                    base_fcf, w, g, terminal_growth, n_years, total_shares
                )
                matrix[i, j] = val if val is not None else np.nan

        add_paragraph(document, (
            f'基于基础FCF={base_fcf / 1e8:.2f}亿元、总股本={total_shares / 1e8:.2f}亿股、'
            f'永续增长率={terminal_growth * 100:.1f}%、预测{n_years}年，'
            f'在WACC [{wacc_min * 100:.0f}%, {wacc_max * 100:.0f}%] x '
            f'增长率 [{growth_min * 100:.0f}%, {growth_max * 100:.0f}%] 范围内计算每股价值。'
        ))

        # 标注当前WACC
        current_wacc = wacc_result.get('wacc', None)
        if current_wacc is not None:
            add_paragraph(document, f'当前WACC: {current_wacc * 100:.2f}%')

        # 敏感性矩阵表格（行列互换展示）
        headers = ['WACC \\ 增长率'] + growth_labels
        table_data = []
        for i, wl in enumerate(wacc_labels):
            row = [wl]
            for j in range(growth_steps):
                v = matrix[i, j]
                row.append(f'{v:.2f}' if not np.isnan(v) else 'N/A')
            table_data.append(row)
        add_table_data(document, headers, table_data)

        # 寻找基准值
        base_wacc = wacc_result.get('wacc', 0.10)
        base_growth = dcf_cfg.get('terminal_growth', 0.03)
        base_value = _compute_dcf_per_share(
            base_fcf, base_wacc, base_growth, terminal_growth, n_years, total_shares
        )
        if base_value is not None:
            add_paragraph(document, f'基准情景(WACC={base_wacc * 100:.2f}%, g={base_growth * 100:.1f}%): 每股价值 {base_value:.2f}元')

    else:
        add_paragraph(document, (
            '缺少基础FCF或总股本数据，无法计算DCF敏感性矩阵。'
            '请确保第四章DCF估值已成功运行。'
        ))
        # 填充占位矩阵
        matrix = np.zeros((wacc_steps, growth_steps))

    # ------------------------------------------------------------------
    # 6.2  热力图
    # ------------------------------------------------------------------
    add_title(document, '6.2 敏感性热力图', level=2)

    heatmap_path = os.path.join(IMAGES_DIR, '06_sensitivity_heatmap.png')
    try:
        center_val = None
        if not np.all(np.isnan(matrix)):
            center_val = float(np.nanmedian(matrix))
        generate_heatmap(
            matrix, wacc_labels, growth_labels,
            heatmap_path,
            title=f'{context["stock_name"]} WACC-增长率敏感性矩阵',
            center=center_val,
            fmt='.2f',
        )
        if os.path.exists(heatmap_path):
            add_paragraph(document, '图表6.1 WACC x 增长率敏感性热力图')
            add_image(document, heatmap_path)
    except Exception as e:
        print(f"  生成热力图失败: {e}")

    # ------------------------------------------------------------------
    # 6.3  时间窗口分析: 20/60/120/250日历史PE分位数
    # ------------------------------------------------------------------
    add_title(document, '6.3 时间窗口PE分位数分析', level=2)
    add_paragraph(document, (
        '在不同时间窗口（20/60/120/250个交易日）下计算当前PE的历史百分位，'
        '评估不同观察期的估值水平。'
    ))

    windows = [20, 60, 120, 250]
    if historical_valuation is not None and not historical_valuation.empty:
        pe_all = historical_valuation['pe_ttm'].dropna()
        pe_all = pe_all[pe_all > 0]
        current_pe = daily_basic.get('pe_ttm', None)

        if not pe_all.empty and current_pe is not None and current_pe > 0:
            headers_win = ['时间窗口', '数据量', '窗口PE均值', '窗口PE中位数', '当前PE', '百分位']
            data_win = []
            for win in windows:
                if len(pe_all) >= win:
                    window_pe = pe_all.iloc[-win:]
                    pct = float((window_pe < current_pe).sum() / len(window_pe) * 100)
                    data_win.append([
                        f'{win}日',
                        str(len(window_pe)),
                        f'{window_pe.mean():.2f}',
                        f'{window_pe.median():.2f}',
                        f'{current_pe:.2f}',
                        f'{pct:.1f}%',
                    ])
                else:
                    data_win.append([f'{win}日', '数据不足', '-', '-', '-', '-'])
            add_table_data(document, headers_win, data_win)

            # 时间窗口柱状图
            try:
                valid_windows = [w for i, w in enumerate(windows) if len(pe_all) >= w]
                valid_pcts = []
                for win in valid_windows:
                    window_pe = pe_all.iloc[-win:]
                    pct = float((window_pe < current_pe).sum() / len(window_pe) * 100)
                    valid_pcts.append(pct)

                if valid_windows:
                    fig, ax = plt.subplots(figsize=(10, 5))
                    x_labels = [f'{w}日' for w in valid_windows]
                    colors = ['#27ae60' if p < 40 else '#f39c12' if p < 60 else '#e74c3c' for p in valid_pcts]
                    bars = ax.bar(x_labels, valid_pcts, color=colors, alpha=0.8, edgecolor='white')
                    ax.axhline(y=50, color='black', linestyle='--', linewidth=1, label='50%分位线')
                    ax.set_ylabel('PE百分位(%)', fontproperties=font_prop, fontsize=12)
                    ax.set_title(f'{context["stock_name"]} 不同时间窗口PE百分位', fontproperties=font_prop, fontsize=14, fontweight='bold')
                    ax.legend(prop=font_prop)
                    ax.grid(True, alpha=0.3, axis='y')
                    for bar, pct_val in zip(bars, valid_pcts):
                        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                                f'{pct_val:.1f}%', ha='center', va='bottom', fontproperties=font_prop, fontsize=10)
                    for label in ax.get_xticklabels():
                        label.set_fontproperties(font_prop)
                    for label in ax.get_yticklabels():
                        label.set_fontproperties(font_prop)
                    plt.tight_layout()
                    win_chart = os.path.join(IMAGES_DIR, '06_time_window_pe.png')
                    plt.savefig(win_chart, dpi=150, bbox_inches='tight')
                    plt.close()
                    if os.path.exists(win_chart):
                        add_paragraph(document, '图表6.2 不同时间窗口PE百分位对比')
                        add_image(document, win_chart)
            except Exception as e:
                print(f"  生成时间窗口图失败: {e}")
        else:
            add_paragraph(document, '历史PE数据不足或当前PE不可用，跳过时间窗口分析。')
    else:
        add_paragraph(document, '历史估值数据不足，跳过时间窗口分析。')

    # ------------------------------------------------------------------
    # 6.4  折价敏感性: 不同折价率(0.7-1.0)下的目标价
    # ------------------------------------------------------------------
    add_title(document, '6.4 折价敏感性分析', level=2)
    add_paragraph(document, (
        '基于DCF估值结果，分析不同折价率（0.70-1.00）对目标价的影响。'
        '折价率反映了流动性折扣、控制权溢价等因素。'
    ))

    dcf_result = results.get('dcf_result', {})
    intrinsic_value = dcf_result.get('per_share_value', None)
    if intrinsic_value is None or intrinsic_value <= 0:
        # 尝试从 matrix 中取中位数
        if not np.all(np.isnan(matrix)):
            intrinsic_value = float(np.nanmedian(matrix))

    discount_rates = np.arange(0.70, 1.01, 0.05)
    headers_disc = ['折价率', '折价因子', '目标价(元)']
    data_disc = []

    if intrinsic_value is not None and intrinsic_value > 0:
        add_paragraph(document, f'基础内在价值: {intrinsic_value:.2f}元/股')
        for dr in discount_rates:
            target_price = intrinsic_value * dr
            data_disc.append([
                f'{dr:.2f}',
                f'{dr:.0%}',
                f'{target_price:.2f}',
            ])
        add_table_data(document, headers_disc, data_disc)

        add_paragraph(document, '')
        current_price = daily_basic.get('close', 0) or 0
        if current_price > 0:
            implied_discount = current_price / intrinsic_value
            add_paragraph(document, f'当前价格 {current_price:.2f}元 对应隐含折价率: {implied_discount:.2%}')
    else:
        add_paragraph(document, 'DCF内在价值不可用，无法进行折价敏感性分析。')

    # ------------------------------------------------------------------
    # 6.5  波动率敏感性: 不同波动率对估值的影响
    # ------------------------------------------------------------------
    add_title(document, '6.5 波动率敏感性分析', level=2)
    add_paragraph(document, (
        '分析不同年化波动率假设对WACC（通过CAPM的Beta调整）及最终估值的影响。'
        '波动率越高，系统性风险越大，要求回报率越高，估值越低。'
    ))

    beta = wacc_result.get('beta', 1.0)
    cost_of_equity = wacc_result.get('cost_of_equity', 0.08)
    wacc_base = wacc_result.get('wacc', 0.10)
    equity_ratio = 0.7
    debt_ratio = 0.3
    risk_free = config.get('wacc', {}).get('risk_free_rate', 0.0185)
    market_premium = config.get('wacc', {}).get('market_risk_premium', 0.0615)

    vol_range = np.arange(0.15, 0.55, 0.05)
    headers_vol = ['年化波动率', '隐含Beta', '权益成本', 'WACC估算']
    data_vol = []

    for vol in vol_range:
        # 简化：波动率与Beta近似线性关系（Beta = 波动率 / 市场波动率）
        market_vol = 0.20  # 假设市场波动率20%
        implied_beta = max(0.3, vol / market_vol)
        implied_coe = risk_free + implied_beta * market_premium
        implied_wacc = implied_coe * equity_ratio + risk_free * 1.5 * debt_ratio * (1 - 0.25)
        data_vol.append([
            f'{vol * 100:.0f}%',
            f'{implied_beta:.2f}',
            f'{implied_coe * 100:.2f}%',
            f'{implied_wacc * 100:.2f}%',
        ])
    add_table_data(document, headers_vol, data_vol)

    add_paragraph(document, '')
    add_paragraph(document, f'当前Beta: {beta:.2f}, 当前WACC: {wacc_base * 100:.2f}%')

    # ------------------------------------------------------------------
    # 6.6  龙卷风图
    # ------------------------------------------------------------------
    add_title(document, '6.6 参数敏感性龙卷风图', level=2)
    add_paragraph(document, (
        '龙卷风图展示各关键参数变化对每股估值的影响幅度，帮助识别最敏感的风险因子。'
    ))

    tornado_path = os.path.join(IMAGES_DIR, '06_tornado_chart.png')

    factors = []
    base_wacc_val = wacc_base
    base_growth_val = terminal_growth

    if base_fcf is not None and total_shares is not None and total_shares > 0:
        base_dcf = _compute_dcf_per_share(
            base_fcf, base_wacc_val, base_growth_val, terminal_growth, n_years, total_shares
        )
        if base_dcf is not None:
            # WACC +/- 2%
            for pct_change, label in [(-0.02, 'WACC -2%'), (0.02, 'WACC +2%')]:
                pass  # compute below

            wacc_low = _compute_dcf_per_share(
                base_fcf, base_wacc_val - 0.02, base_growth_val, terminal_growth, n_years, total_shares
            )
            wacc_high = _compute_dcf_per_share(
                base_fcf, base_wacc_val + 0.02, base_growth_val, terminal_growth, n_years, total_shares
            )
            if wacc_low is not None and wacc_high is not None:
                factors.append({
                    'name': 'WACC (+/-2%)',
                    'low': wacc_low,
                    'high': wacc_high,
                    'range': abs(wacc_high - wacc_low),
                })

            # 增长率 +/- 3%
            growth_low = _compute_dcf_per_share(
                base_fcf, base_wacc_val, base_growth_val - 0.03, terminal_growth, n_years, total_shares
            )
            growth_high = _compute_dcf_per_share(
                base_fcf, base_wacc_val, base_growth_val + 0.03, terminal_growth, n_years, total_shares
            )
            if growth_low is not None and growth_high is not None:
                factors.append({
                    'name': '增长率 (+/-3%)',
                    'low': growth_low,
                    'high': growth_high,
                    'range': abs(growth_high - growth_low),
                })

            # 永续增长率 +/- 1%
            tg_low = _compute_dcf_per_share(
                base_fcf, base_wacc_val, base_growth_val, terminal_growth - 0.01, n_years, total_shares
            )
            tg_high = _compute_dcf_per_share(
                base_fcf, base_wacc_val, base_growth_val, terminal_growth + 0.01, n_years, total_shares
            )
            if tg_low is not None and tg_high is not None:
                factors.append({
                    'name': '永续增长率 (+/-1%)',
                    'low': tg_low,
                    'high': tg_high,
                    'range': abs(tg_high - tg_low),
                })

            # FCF +/- 10%
            fcf_low = _compute_dcf_per_share(
                base_fcf * 0.9, base_wacc_val, base_growth_val, terminal_growth, n_years, total_shares
            )
            fcf_high = _compute_dcf_per_share(
                base_fcf * 1.1, base_wacc_val, base_growth_val, terminal_growth, n_years, total_shares
            )
            if fcf_low is not None and fcf_high is not None:
                factors.append({
                    'name': 'FCF (+/-10%)',
                    'low': fcf_low,
                    'high': fcf_high,
                    'range': abs(fcf_high - fcf_low),
                })

        if factors:
            try:
                generate_tornado_chart(
                    factors, base_dcf, tornado_path,
                    title=f'{context["stock_name"]} 参数敏感性龙卷风图',
                )
                if os.path.exists(tornado_path):
                    add_paragraph(document, '图表6.3 参数敏感性龙卷风图')
                    add_image(document, tornado_path)
            except Exception as e:
                print(f"  生成龙卷风图失败: {e}")

            add_paragraph(document, '')
            add_paragraph(document, '龙卷风图分析结论:')
            for f in sorted(factors, key=lambda x: x['range'], reverse=True):
                add_paragraph(document, (
                    f'- {f["name"]}: 影响范围 {f["range"]:.2f}元 '
                    f'({f["low"]:.2f} ~ {f["high"]:.2f}元)'
                ))
        else:
            add_paragraph(document, '无法计算敏感性因子（DCF参数不完整）。')
    else:
        add_paragraph(document, '缺少基础数据，无法生成龙卷风图。')

    add_section_break(document)

    # ==================== 保存结果 ====================
    context['results']['sensitivity_matrix'] = matrix.tolist()

    return context
