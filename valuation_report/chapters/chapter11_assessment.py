#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第十一章 - 综合评估

功能：加权估值汇总、五维风险雷达图、反向测算、投资建议、风险提示清单
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy.optimize import brentq

from module_utils import add_title, add_paragraph, add_table_data, add_image, add_section_break


def _generate_implied_analysis_chart(growth_range, values_growth, wacc_range, values_wacc,
                                      current_price, save_path, stock_name='', font_prop=None):
    """生成反向测算分析图"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))

    # 增长率敏感性
    ax1.plot([g * 100 for g in growth_range], values_growth, 'b-', linewidth=2)
    ax1.axhline(y=current_price, color='red', linestyle='--', linewidth=1.5,
                label=f'当前价格 {current_price:.2f}元')
    ax1.set_xlabel('永续增长率(%)', fontproperties=font_prop, fontsize=12)
    ax1.set_ylabel('DCF每股价值(元)', fontproperties=font_prop, fontsize=12)
    ax1.set_title('隐含增长率测算', fontproperties=font_prop, fontsize=13, fontweight='bold')
    ax1.legend(prop=font_prop, fontsize=10)
    ax1.grid(True, alpha=0.3)
    for label in ax1.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax1.get_yticklabels():
        label.set_fontproperties(font_prop)

    # WACC敏感性
    ax2.plot([w * 100 for w in wacc_range], values_wacc, 'b-', linewidth=2)
    ax2.axhline(y=current_price, color='red', linestyle='--', linewidth=1.5,
                label=f'当前价格 {current_price:.2f}元')
    ax2.set_xlabel('WACC(%)', fontproperties=font_prop, fontsize=12)
    ax2.set_ylabel('DCF每股价值(元)', fontproperties=font_prop, fontsize=12)
    ax2.set_title('隐含WACC测算', fontproperties=font_prop, fontsize=13, fontweight='bold')
    ax2.legend(prop=font_prop, fontsize=10)
    ax2.grid(True, alpha=0.3)
    for label in ax2.get_xticklabels():
        label.set_fontproperties(font_prop)
    for label in ax2.get_yticklabels():
        label.set_fontproperties(font_prop)

    plt.suptitle(f'{stock_name} 反向测算分析', fontproperties=font_prop, fontsize=15, fontweight='bold', y=1.02)
    plt.tight_layout()
    plt.savefig(save_path, dpi=150, bbox_inches='tight')
    plt.close()


def generate_chapter(context):
    """
    生成第十一章：综合评估

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

    daily_basic = context.get('daily_basic', {})
    wacc_result = context.get('wacc_result', {})
    price_data = context.get('price_data', pd.DataFrame())
    historical_valuation = context.get('historical_valuation', pd.DataFrame())
    peer_valuation = context.get('peer_valuation', pd.DataFrame())
    financial_indicators = context.get('financial_indicators', pd.DataFrame())
    financial_statements = context.get('financial_statements', {})
    results = context.get('results', {})

    current_price = float(daily_basic.get('close', 0))
    pe_ttm = float(daily_basic.get('pe_ttm', 0))
    pb = float(daily_basic.get('pb', 0))
    total_mv = float(daily_basic.get('total_mv', 0))
    wacc = float(wacc_result.get('wacc', 0.08)) if wacc_result else 0.08

    # 获取各章节结果
    dcf_summary = results.get('dcf_summary', {})
    relative_valuation = results.get('relative_valuation', {})
    mc_result = results.get('mc_result', {})
    var_result = results.get('var_result', {})
    sensitivity_matrix = results.get('sensitivity_matrix', {})

    # dcf_summary可能是list（chapter04输出）或dict
    if isinstance(dcf_summary, list) and dcf_summary:
        dcf_per_share = float(dcf_summary[0].get('per_share', 0)) if isinstance(dcf_summary[0], dict) else 0
    elif isinstance(dcf_summary, dict):
        dcf_per_share = float(dcf_summary.get('per_share', 0))
    else:
        dcf_per_share = float(results.get('dcf_avg_per_share', 0))
    relative_per_share = float(relative_valuation.get('target_price', 0)) if relative_valuation else 0
    mc_median_price = float(mc_result.get('median_price', 0)) if mc_result else 0

    # ==================== 十一、综合评估 ====================
    add_title(document, '十一、综合评估', level=1)
    add_paragraph(document,
        '本章综合前面各章的分析结果，通过加权估值汇总、五维风险评估、反向测算等方法，'
        '给出最终的投资建议和风险提示。')

    # ==================== 11.1 加权估值汇总 ====================
    add_title(document, '11.1 加权估值汇总', level=2)
    add_paragraph(document,
        '将DCF绝对估值、相对估值、蒙特卡洛模拟三种方法的估值结果按权重加总，'
        '得到综合目标价格。权重分配考虑了各方法的可靠性和适用性。')

    # 权重设置
    weights = {'DCF': 0.40, 'Relative': 0.30, 'MC': 0.30}

    # 收集有效估值
    valuations = {}
    if dcf_per_share > 0:
        valuations['DCF绝对估值'] = dcf_per_share
    if relative_per_share > 0:
        valuations['相对估值'] = relative_per_share
    if mc_median_price > 0:
        valuations['蒙特卡洛模拟'] = mc_median_price

    # 如果某种方法无结果，重新分配权重
    active_weights = {}
    weight_map = {'DCF绝对估值': 'DCF', '相对估值': 'Relative', '蒙特卡洛模拟': 'MC'}
    for key, val in valuations.items():
        wk = weight_map.get(key, 'DCF')
        active_weights[key] = weights[wk]

    # 归一化权重
    total_weight = sum(active_weights.values())
    if total_weight > 0:
        for key in active_weights:
            active_weights[key] /= total_weight
    else:
        # 均分
        n = len(valuations)
        for key in valuations:
            active_weights[key] = 1.0 / n if n > 0 else 0

    # 计算加权目标价
    weighted_target = sum(valuations[key] * active_weights.get(key, 0) for key in valuations)

    # 构建估值汇总表
    summary_headers = ['估值方法', '估值结果(元/股)', '权重', '加权贡献(元/股)']
    summary_data = []
    for key in valuations:
        summary_data.append([
            key,
            f'{valuations[key]:.2f}',
            f'{active_weights.get(key, 0):.0%}',
            f'{valuations[key] * active_weights.get(key, 0):.2f}'
        ])
    summary_data.append([
        '加权目标价',
        f'{weighted_target:.2f}',
        '100%',
        f'{weighted_target:.2f}'
    ])
    add_table_data(document, summary_headers, summary_data)

    add_paragraph(document, '')
    if weighted_target > 0 and current_price > 0:
        upside = (weighted_target - current_price) / current_price * 100
        add_paragraph(document,
            f'综合加权目标价格为{weighted_target:.2f}元/股，当前价格{current_price:.2f}元/股，'
            f'{"上涨" if upside > 0 else "下跌"}空间{abs(upside):.2f}%。', bold=True)
    elif weighted_target > 0:
        add_paragraph(document, f'综合加权目标价格为{weighted_target:.2f}元/股。', bold=True)
    else:
        add_paragraph(document, '各估值方法均未产出有效结果，无法计算综合目标价格。', bold=True)

    # 估值区间
    if valuations:
        vals = list(valuations.values())
        low = min(vals)
        high = max(vals)
        add_paragraph(document,
            f'估值区间：{low:.2f}元 ~ {high:.2f}元（各方法最低至最高估值）')

    # ==================== 11.2 五维风险雷达图 ====================
    add_paragraph(document, '')
    add_title(document, '11.2 五维风险雷达图', level=2)
    add_paragraph(document,
        '从估值风险、财务风险、市场风险、流动性风险、行业风险五个维度进行综合评分，'
        '每项得分0-100分，分数越高表示该维度风险越低（状况越好）。')

    # 计算各维度得分
    scores = {}
    labels = {}

    # 1. 估值风险 (基于PE百分位)
    _pe_raw = relative_valuation.get('pe_percentile', 50) if relative_valuation else 50
    pe_percentile = float(_pe_raw) if _pe_raw is not None else 50
    if pe_percentile <= 20:
        scores['valuation'] = 90
    elif pe_percentile <= 40:
        scores['valuation'] = 75
    elif pe_percentile <= 60:
        scores['valuation'] = 60
    elif pe_percentile <= 80:
        scores['valuation'] = 40
    else:
        scores['valuation'] = 20
    labels['valuation'] = f'估值风险({scores["valuation"]}分)'

    # 2. 财务风险 (基于ROE和资产负债率)
    roe = 0
    debt_ratio = 0
    if financial_indicators is not None and not financial_indicators.empty:
        if 'roe' in financial_indicators.columns:
            roe_vals = financial_indicators['roe'].dropna()
            if not roe_vals.empty:
                roe = float(roe_vals.iloc[0])
        if 'debt_to_assets' in financial_indicators.columns:
            debt_vals = financial_indicators['debt_to_assets'].dropna()
            if not debt_vals.empty:
                debt_ratio = float(debt_vals.iloc[0])
                # Tushare fina_indicator返回百分比值(0-100)，需转为小数(0-1)
                if debt_ratio > 1:
                    debt_ratio = debt_ratio / 100

    # 从资产负债表获取
    if debt_ratio == 0:
        bs = financial_statements.get('balancesheet', pd.DataFrame())
        if bs is not None and not bs.empty:
            if 'total_liab' in bs.columns and 'total_assets' in bs.columns:
                total_liab = float(bs['total_liab'].dropna().iloc[0]) if not bs['total_liab'].dropna().empty else 0
                total_assets = float(bs['total_assets'].dropna().iloc[0]) if not bs['total_assets'].dropna().empty else 1
                if total_assets > 0:
                    debt_ratio = total_liab / total_assets

    # ROE得分 (ROE越高越好)
    if roe >= 20:
        roe_score = 90
    elif roe >= 15:
        roe_score = 75
    elif roe >= 10:
        roe_score = 60
    elif roe >= 5:
        roe_score = 40
    else:
        roe_score = 20

    # 负债率得分 (越低越好)
    if debt_ratio <= 0.30:
        debt_score = 90
    elif debt_ratio <= 0.50:
        debt_score = 75
    elif debt_ratio <= 0.65:
        debt_score = 55
    elif debt_ratio <= 0.80:
        debt_score = 35
    else:
        debt_score = 15

    scores['financial'] = int(roe_score * 0.6 + debt_score * 0.4)
    labels['financial'] = f'财务风险({scores["financial"]}分)'

    # 3. 市场风险 (基于波动率)
    if price_data is not None and not price_data.empty and 'close' in price_data.columns:
        prices = price_data['close'].astype(float).sort_index()
        daily_returns = prices.pct_change().dropna()
        if len(daily_returns) > 20:
            annual_vol = float(daily_returns.tail(60).std() * np.sqrt(252))
        else:
            annual_vol = 0.30
    else:
        annual_vol = 0.30

    if annual_vol <= 0.20:
        scores['market'] = 85
    elif annual_vol <= 0.30:
        scores['market'] = 70
    elif annual_vol <= 0.40:
        scores['market'] = 55
    elif annual_vol <= 0.55:
        scores['market'] = 35
    else:
        scores['market'] = 15
    labels['market'] = f'市场风险({scores["market"]}分)'

    # 4. 流动性风险 (基于成交额)
    if price_data is not None and not price_data.empty and 'amount' in price_data.columns:
        avg_amount = float(price_data['amount'].tail(20).mean())
        if avg_amount >= 500000:     # >5亿
            scores['liquidity'] = 90
        elif avg_amount >= 200000:   # >2亿
            scores['liquidity'] = 75
        elif avg_amount >= 100000:   # >1亿
            scores['liquidity'] = 60
        elif avg_amount >= 50000:    # >5000万
            scores['liquidity'] = 40
        else:
            scores['liquidity'] = 20
    else:
        scores['liquidity'] = 50
    labels['liquidity'] = f'流动性风险({scores["liquidity"]}分)'

    # 5. 行业风险 (基于行业PE位置)
    if peer_valuation is not None and not peer_valuation.empty and 'pe_ttm' in peer_valuation.columns:
        peer_pe = peer_valuation['pe_ttm'].dropna()
        peer_pe = peer_pe[peer_pe > 0]
        if len(peer_pe) >= 3 and pe_ttm > 0:
            pe_position = float((peer_pe < pe_ttm).mean()) * 100
            if pe_position <= 25:
                scores['industry'] = 85
            elif pe_position <= 50:
                scores['industry'] = 70
            elif pe_position <= 75:
                scores['industry'] = 50
            else:
                scores['industry'] = 30
        else:
            scores['industry'] = 50
    else:
        scores['industry'] = 50
    labels['industry'] = f'行业风险({scores["industry"]}分)'

    # 输出得分详情
    score_detail_headers = ['维度', '得分', '评分依据']
    score_detail_data = [
        ['估值风险', str(scores['valuation']),
         f'PE百分位{pe_percentile:.0f}%'],
        ['财务风险', str(scores['financial']),
         f'ROE={roe:.2f}%, 资产负债率={debt_ratio:.2%}'],
        ['市场风险', str(scores['market']),
         f'年化波动率={annual_vol:.2%}'],
        ['流动性风险', str(scores['liquidity']),
         f'近20日均成交额={avg_amount / 10:.0f}万元' if price_data is not None and not price_data.empty and 'amount' in price_data.columns else '数据不足'],
        ['行业风险', str(scores['industry']),
         f'PE在同行中的位置'],
    ]
    add_table_data(document, score_detail_headers, score_detail_data)

    # 生成雷达图
    add_paragraph(document, '')
    radar_path = os.path.join(IMAGES_DIR, f'{stock_code.replace(".", "_")}_radar.png')
    try:
        from module_utils import generate_radar_chart
        generate_radar_chart(scores, labels, radar_path, title=f'{stock_name} 五维风险雷达图')
        add_image(document, radar_path)
    except Exception as e:
        add_paragraph(document, f'[雷达图生成失败: {e}]')

    # 总体评分
    overall_score = np.mean(list(scores.values()))
    add_paragraph(document, '')
    add_paragraph(document, f'综合风险评分：{overall_score:.1f}分（满分100分）', bold=True)

    if overall_score >= 75:
        risk_assessment = '低风险'
        risk_desc = '各维度风险均处于较低水平，投资安全性较高。'
    elif overall_score >= 60:
        risk_assessment = '中低风险'
        risk_desc = '整体风险可控，部分维度需关注。'
    elif overall_score >= 45:
        risk_assessment = '中等风险'
        risk_desc = '存在一定风险因素，需审慎评估。'
    elif overall_score >= 30:
        risk_assessment = '中高风险'
        risk_desc = '多个维度存在风险信号，建议降低仓位。'
    else:
        risk_assessment = '高风险'
        risk_desc = '多个维度风险较高，建议谨慎投资或规避。'

    add_paragraph(document, f'风险等级：{risk_assessment} - {risk_desc}')

    # 找出最弱的维度
    weakest = min(scores, key=scores.get)
    weakest_labels = {
        'valuation': '估值风险',
        'financial': '财务风险',
        'market': '市场风险',
        'liquidity': '流动性风险',
        'industry': '行业风险',
    }
    add_paragraph(document,
        f'最需关注的维度：{weakest_labels.get(weakest, weakest)}（{scores[weakest]}分），'
        f'建议重点关注该维度相关风险因素。')

    # ==================== 11.3 反向测算 ====================
    add_paragraph(document, '')
    add_title(document, '11.3 反向测算', level=2)
    add_paragraph(document,
        '反向测算从当前市场价格出发，计算隐含的增长率和WACC。'
        '通过与基准参数对比，判断市场定价的合理性。')

    implied_growth = None
    implied_wacc_val = None

    if dcf_per_share > 0 and current_price > 0 and wacc > 0:
        # 从DCF结果获取基本参数
        fcf_list = dcf_summary.get('projected_fcf', [])
        terminal_growth_base = float(dcf_summary.get('terminal_growth', 0.03))
        shares = float(dcf_summary.get('total_shares', daily_basic.get('total_share', 1)))

        # 使用简化的Gordon Growth Model进行反向测算
        # value = FCF1 / (WACC - g)
        # 如果有DCF的详细参数，用敏感性矩阵的方式反推

        # 隐含增长率测算：固定WACC，求解g使得DCF = current_price
        try:
            # 获取基准FCF
            base_fcf = float(dcf_summary.get('base_fcf', fcf_list[0] if fcf_list else 0))
            if base_fcf <= 0:
                base_fcf = dcf_per_share * shares * wacc  # 反推

            # 使用简化模型：V = FCF * (1+g) / (WACC - g)
            # 每股价值 = V / shares
            def dcf_value_for_growth(g, fcf=base_fcf, w=wacc, n_shares=shares):
                if w <= g:
                    return 1e10
                return fcf * (1 + g) / (w - g) / n_shares

            growth_range = np.linspace(-0.05, 0.15, 200)
            values_growth = [dcf_value_for_growth(g) for g in growth_range]

            # 检查当前价格是否在范围内
            min_val = min(values_growth)
            max_val = max(values_growth)
            if min_val <= current_price <= max_val:
                implied_growth = brentq(lambda g: dcf_value_for_growth(g) - current_price, -0.05, wacc - 0.001)
            elif current_price < min_val:
                implied_growth = growth_range[0]
            else:
                implied_growth = growth_range[-1]

        except Exception:
            implied_growth = None

        # 隐含WACC测算：固定g，求解w使得DCF = current_price
        try:
            def dcf_value_for_wacc(w, fcf=base_fcf, g=terminal_growth_base, n_shares=shares):
                if w <= g:
                    return 1e10
                return fcf * (1 + g) / (w - g) / n_shares

            wacc_range = np.linspace(0.03, 0.20, 200)
            values_wacc = [dcf_value_for_wacc(w) for w in wacc_range]

            min_val_w = min(values_wacc)
            max_val_w = max(values_wacc)
            if min_val_w <= current_price <= max_val_w:
                implied_wacc_val = brentq(lambda w: dcf_value_for_wacc(w) - current_price, terminal_growth_base + 0.001, 0.20)
            elif current_price < min_val_w:
                implied_wacc_val = wacc_range[-1]
            else:
                implied_wacc_val = wacc_range[0]

        except Exception:
            implied_wacc_val = None

        # 生成反向测算图表
        if len(growth_range) > 0 and len(values_growth) > 0:
            implied_chart_path = os.path.join(IMAGES_DIR,
                                               f'{stock_code.replace(".", "_")}_implied_analysis.png')
            try:
                _generate_implied_analysis_chart(
                    growth_range, values_growth, wacc_range, values_wacc,
                    current_price, implied_chart_path, stock_name, font_prop)
                add_image(document, implied_chart_path)
            except Exception as e:
                add_paragraph(document, f'[反向测算图表生成失败: {e}]')

    # 输出反向测算结果
    if implied_growth is not None or implied_wacc_val is not None:
        implied_headers = ['反向测算指标', '隐含值', '基准值', '差异', '含义']
        implied_data = []

        if implied_growth is not None:
            growth_diff = implied_growth - terminal_growth_base
            if implied_growth > terminal_growth_base + 0.03:
                growth_meaning = '市场定价偏乐观，隐含增速高于合理水平'
            elif implied_growth < terminal_growth_base - 0.03:
                growth_meaning = '市场定价偏悲观，隐含增速低于合理水平'
            else:
                growth_meaning = '市场定价隐含增速与基准接近，定价合理'

            implied_data.append([
                '隐含永续增长率',
                f'{implied_growth:.2%}',
                f'{terminal_growth_base:.2%}',
                f'{growth_diff:+.2%}',
                growth_meaning,
            ])

        if implied_wacc_val is not None:
            wacc_diff = implied_wacc_val - wacc
            if implied_wacc_val > wacc + 0.02:
                wacc_meaning = '市场要求更高的风险补偿，定价偏悲观'
            elif implied_wacc_val < wacc - 0.02:
                wacc_meaning = '市场要求较低的风险补偿，定价偏乐观'
            else:
                wacc_meaning = '市场风险偏好与基准接近，定价合理'

            implied_data.append([
                '隐含WACC',
                f'{implied_wacc_val:.2%}',
                f'{wacc:.2%}',
                f'{wacc_diff:+.2%}',
                wacc_meaning,
            ])

        add_table_data(document, implied_headers, implied_data)
    else:
        add_paragraph(document, '由于DCF参数不完整，无法进行反向测算。')

    # ==================== 11.4 投资建议 ====================
    add_paragraph(document, '')
    add_title(document, '11.4 投资建议', level=2)

    if weighted_target > 0 and current_price > 0:
        upside_ratio = weighted_target / current_price

        if upside_ratio > 1.15:
            advice = '买入'
            advice_color = '看多'
            threshold_upper = weighted_target
            threshold_lower = current_price * 1.15
            add_paragraph(document,
                f'投资建议：{advice}', bold=True)
            add_paragraph(document,
                f'综合目标价格{weighted_target:.2f}元/股，较当前价格{current_price:.2f}元/股'
                f'存在{(upside_ratio - 1) * 100:.1f}%的上涨空间，'
                f'高于15%的买入阈值，建议买入。')
            add_paragraph(document,
                f'合理价格区间：{threshold_lower:.2f}元 ~ {threshold_upper:.2f}元')

        elif upside_ratio >= 0.85:
            advice = '持有'
            advice_color = '中性'
            add_paragraph(document,
                f'投资建议：{advice}', bold=True)
            add_paragraph(document,
                f'综合目标价格{weighted_target:.2f}元/股，与当前价格{current_price:.2f}元/股'
                f'接近（偏离度{(upside_ratio - 1) * 100:+.1f}%），'
                f'估值处于合理区间，建议持有。')

        else:
            advice = '卖出'
            advice_color = '看空'
            add_paragraph(document,
                f'投资建议：{advice}', bold=True)
            add_paragraph(document,
                f'综合目标价格{weighted_target:.2f}元/股，较当前价格{current_price:.2f}元/股'
                f'存在{(1 - upside_ratio) * 100:.1f}%的高估空间，'
                f'低于85%的卖出阈值，建议卖出。')

        # 价格区间分析
        add_paragraph(document, '')
        add_paragraph(document, '估值分布区间：', bold=True)
        price_range_headers = ['估值方法', '估值(元/股)', '较当前价(%)']
        price_range_data = []
        for key in valuations:
            pct = (valuations[key] - current_price) / current_price * 100
            price_range_data.append([
                key,
                f'{valuations[key]:.2f}',
                f'{pct:+.2f}'
            ])
        price_range_data.append([
            '加权目标价',
            f'{weighted_target:.2f}',
            f'{(weighted_target - current_price) / current_price * 100:+.2f}'
        ])
        add_table_data(document, price_range_headers, price_range_data)

        # 风险收益比
        if var_result:
            var_95_val = float(var_result.get('var_95', 0))
            max_dd = float(var_result.get('max_drawdown', 0))
            if var_95_val != 0 or max_dd != 0:
                add_paragraph(document, '')
                add_paragraph(document, '风险收益特征：', bold=True)
                potential_return = (weighted_target - current_price) / current_price
                risk_reward_headers = ['指标', '数值']
                risk_reward_data = [
                    ['潜在上涨空间', f'{potential_return:.2%}'],
                    ['95% VaR(日)', f'{var_95_val:.2%}'],
                    ['最大回撤(历史)', f'{max_dd:.2%}'],
                    ['风险收益比', f'{abs(potential_return / abs(max_dd)):.2f}' if max_dd != 0 else 'N/A'],
                ]
                add_table_data(document, risk_reward_headers, risk_reward_data)

    else:
        add_paragraph(document,
            '由于估值数据不充分，无法给出明确投资建议。建议结合基本面和行业分析进行判断。')

    # ==================== 11.5 风险提示清单 ====================
    add_paragraph(document, '')
    add_title(document, '11.5 风险提示清单', level=2)
    add_paragraph(document,
        '投资者在做出投资决策前，应充分了解以下风险因素：')

    risk_items = []

    # 估值风险
    if pe_ttm > 0:
        if pe_percentile > 70:
            risk_items.append(('估值风险', '当前PE处于历史较高百分位，存在估值回调风险。'))
        elif pe_percentile < 30:
            risk_items.append(('估值风险', '当前PE处于历史较低百分位，需关注是否存在基本面恶化导致估值走低。'))

    # 市场风险
    risk_items.append(('市场风险', '宏观经济波动、政策调整、市场情绪变化等系统性风险可能导致股价大幅波动。'))

    # 流动性风险
    risk_items.append(('流动性风险', '在市场极端情况下，可能出现流动性枯竭，导致无法及时卖出。'))

    # 行业风险
    risk_items.append(('行业风险', '行业竞争格局变化、技术迭代、政策监管等因素可能影响公司盈利能力。'))

    # 财务风险
    if debt_ratio > 0.6:
        risk_items.append(('财务风险', f'资产负债率较高（{debt_ratio:.1%}），存在一定偿债压力。'))

    # 模型风险
    risk_items.append(('模型风险',
        '本报告的估值结果基于多项假设和模型参数，实际结果可能与预测存在显著偏差。'
        'DCF模型对增长率和折现率高度敏感，蒙特卡洛模拟基于历史数据外推，'
        '不保证未来收益。'))

    # 信息风险
    risk_items.append(('信息风险',
        '本报告所使用的财务数据和估值数据来源于公开信息（Tushare），'
        '数据的完整性和准确性可能影响分析结论。'))

    # VaR风险提示
    if var_result:
        var_95_v = float(var_result.get('var_95', 0))
        max_dd_v = float(var_result.get('max_drawdown', 0))
        if abs(max_dd_v) > 0.30:
            risk_items.append(('下行风险',
                f'历史最大回撤达{max_dd_v:.2%}，投资者需具备相应的风险承受能力。'))

    # 输出风险提示
    risk_headers = ['序号', '风险类型', '风险说明']
    risk_data = []
    for idx, (rtype, rdesc) in enumerate(risk_items, 1):
        risk_data.append([str(idx), rtype, rdesc])
    add_table_data(document, risk_headers, risk_data)

    add_paragraph(document, '')
    add_paragraph(document,
        '免责声明：本报告仅供投资参考，不构成任何投资建议。投资者应根据自身风险承受能力'
        '独立做出投资决策，投资有风险，入市需谨慎。', bold=True)

    add_section_break(document)
    return context
