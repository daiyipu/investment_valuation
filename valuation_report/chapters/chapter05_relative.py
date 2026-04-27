#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第五章 - 相对估值分析

包含:
- 5.1 PE/PB 3年滚动百分位
- 5.2 同行PE/PB对比表
- 5.3 SW行业指数PE比较
- 5.4 估值水平判定
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
    generate_pe_percentile_chart,
)


def generate_chapter(context):
    """生成第五章：相对估值分析

    Args:
        context: dict with keys:
            stock_code, stock_name, report_date, document, font_prop,
            config, pro (tushare), company_info, sw_industry, peer_codes,
            price_data, daily_basic, market_indices, financial_statements,
            financial_indicators, historical_valuation (DataFrame),
            peer_valuation (DataFrame), wacc_result, IMAGES_DIR, OUTPUT_DIR,
            results (dict)

    Returns:
        context (updated in-place, results['relative_valuation'] populated)
    """
    document = context['document']
    font_prop = context['font_prop']
    pro = context['pro']
    stock_code = context['stock_code']
    stock_name = context['stock_name']
    historical_valuation = context['historical_valuation']
    peer_valuation = context['peer_valuation']
    daily_basic = context['daily_basic']
    sw_industry = context.get('sw_industry', {})
    IMAGES_DIR = context['IMAGES_DIR']

    # ==================== 五、相对估值分析 ====================
    add_title(document, '五、相对估值分析', level=1)
    add_paragraph(document, (
        '本章节通过相对估值法，从历史估值百分位、同行对比、行业指数对比等维度，'
        '综合评估标的股票的估值水平。'
    ))

    # ------------------------------------------------------------------
    # 5.1  PE/PB 3年滚动百分位
    # ------------------------------------------------------------------
    add_title(document, '5.1 PE/PB 3年滚动百分位', level=2)
    add_paragraph(document, (
        '基于过去3年（约750个交易日）的历史PE(TTM)和PB数据，计算当前估值在历史区间中的百分位位置。'
        '百分位越低说明估值越便宜，越高说明估值越贵。'
    ))

    pe_percentile = None
    pb_percentile = None
    pe_stats = {}
    pb_stats = {}

    if historical_valuation is not None and not historical_valuation.empty:
        hv = historical_valuation.copy()

        # ---- PE 百分位 ----
        pe_series = hv['pe_ttm'].dropna()
        if not pe_series.empty and pe_series.max() > 0:
            # 过滤极端值（负PE和超高PE）
            pe_valid = pe_series[(pe_series > 0) & (pe_series < 500)]
            if not pe_valid.empty:
                current_pe = daily_basic.get('pe_ttm', None)
                if current_pe is not None and current_pe > 0:
                    pe_stats['min'] = float(pe_valid.min())
                    pe_stats['p25'] = float(pe_valid.quantile(0.25))
                    pe_stats['p50'] = float(pe_valid.quantile(0.50))
                    pe_stats['p75'] = float(pe_valid.quantile(0.75))
                    pe_stats['max'] = float(pe_valid.max())
                    pe_stats['current'] = float(current_pe)
                    pe_percentile = float((pe_valid < current_pe).sum() / len(pe_valid) * 100)
                    pe_stats['percentile'] = pe_percentile

        # ---- PB 百分位 ----
        pb_series = hv['pb'].dropna()
        if not pb_series.empty and pb_series.max() > 0:
            pb_valid = pb_series[(pb_series > 0) & (pb_series < 20)]
            if not pb_valid.empty:
                current_pb = daily_basic.get('pb', None)
                if current_pb is not None and current_pb > 0:
                    pb_stats['min'] = float(pb_valid.min())
                    pb_stats['p25'] = float(pb_valid.quantile(0.25))
                    pb_stats['p50'] = float(pb_valid.quantile(0.50))
                    pb_stats['p75'] = float(pb_valid.quantile(0.75))
                    pb_stats['max'] = float(pb_valid.max())
                    pb_stats['current'] = float(current_pb)
                    pb_percentile = float((pb_valid < current_pb).sum() / len(pb_valid) * 100)
                    pb_stats['percentile'] = pb_percentile

        # ---- 输出百分位表格 ----
        if pe_stats and pb_stats:
            headers = ['指标', '当前值', '最小值', '25%分位', '50%分位', '75%分位', '最大值', '当前百分位']
            data = [
                [
                    'PE(TTM)',
                    f"{pe_stats['current']:.2f}",
                    f"{pe_stats['min']:.2f}",
                    f"{pe_stats['p25']:.2f}",
                    f"{pe_stats['p50']:.2f}",
                    f"{pe_stats['p75']:.2f}",
                    f"{pe_stats['max']:.2f}",
                    f"{pe_stats['percentile']:.1f}%",
                ],
                [
                    'PB',
                    f"{pb_stats['current']:.2f}",
                    f"{pb_stats['min']:.2f}",
                    f"{pb_stats['p25']:.2f}",
                    f"{pb_stats['p50']:.2f}",
                    f"{pb_stats['p75']:.2f}",
                    f"{pb_stats['max']:.2f}",
                    f"{pb_stats['percentile']:.1f}%",
                ],
            ]
            add_table_data(document, headers, data)
        else:
            add_paragraph(document, '历史PE/PB数据不完整，无法计算百分位。')

        # ---- PE 百分位趋势图 ----
        if not pe_series.empty:
            pe_series_indexed = pe_valid.reset_index(drop=True)
            chart_path = os.path.join(IMAGES_DIR, '05_pe_percentile.png')
            try:
                generate_pe_percentile_chart(
                    pe_series_indexed, chart_path,
                    stock_name=stock_name, window=min(250, len(pe_series_indexed))
                )
                if os.path.exists(chart_path):
                    add_paragraph(document, f'图表5.1 {stock_name} PE(TTM)历史走势与百分位')
                    add_image(document, chart_path)
            except Exception as e:
                print(f"  生成PE百分位图失败: {e}")

    else:
        add_paragraph(document, '历史估值数据不足，跳过3年滚动百分位分析。')

    # ------------------------------------------------------------------
    # 5.2  同行PE/PB对比表
    # ------------------------------------------------------------------
    add_title(document, '5.2 同行PE/PB对比表', level=2)
    add_paragraph(document, (
        '将标的股票与申万三级行业内的可比公司进行PE(TTM)和PB对比，'
        '计算行业均值、中位数、Q1/Q3和标准差，评估标的在同行中的估值位置。'
    ))

    industry_pe_stats = {}
    industry_pb_stats = {}

    if peer_valuation is not None and not peer_valuation.empty:
        pv = peer_valuation.copy()

        # 过滤异常值
        pv_pe = pv[(pv['pe_ttm'] > 0) & (pv['pe_ttm'] < 500)]
        pv_pb = pv[(pv['pb'] > 0) & (pv['pb'] < 20)]

        if not pv_pe.empty:
            industry_pe_stats = {
                'mean': float(pv_pe['pe_ttm'].mean()),
                'median': float(pv_pe['pe_ttm'].median()),
                'q1': float(pv_pe['pe_ttm'].quantile(0.25)),
                'q3': float(pv_pe['pe_ttm'].quantile(0.75)),
                'std': float(pv_pe['pe_ttm'].std()),
            }
        if not pv_pb.empty:
            industry_pb_stats = {
                'mean': float(pv_pb['pb'].mean()),
                'median': float(pv_pb['pb'].median()),
                'q1': float(pv_pb['pb'].quantile(0.25)),
                'q3': float(pv_pb['pb'].quantile(0.75)),
                'std': float(pv_pb['pb'].std()),
            }

        # 行业统计表
        current_pe = daily_basic.get('pe_ttm', 0) or 0
        current_pb = daily_basic.get('pb', 0) or 0

        if industry_pe_stats and industry_pb_stats:
            headers = ['指标', f'{stock_name}(标的)', '行业均值', '行业中位数',
                        '行业Q1', '行业Q3', '标准差', '偏离均值']
            data = [
                [
                    'PE(TTM)',
                    f"{current_pe:.2f}",
                    f"{industry_pe_stats['mean']:.2f}",
                    f"{industry_pe_stats['median']:.2f}",
                    f"{industry_pe_stats['q1']:.2f}",
                    f"{industry_pe_stats['q3']:.2f}",
                    f"{industry_pe_stats['std']:.2f}",
                    f"{((current_pe - industry_pe_stats['mean']) / industry_pe_stats['mean'] * 100):+.1f}%" if industry_pe_stats['mean'] != 0 else 'N/A',
                ],
                [
                    'PB',
                    f"{current_pb:.2f}",
                    f"{industry_pb_stats['mean']:.2f}",
                    f"{industry_pb_stats['median']:.2f}",
                    f"{industry_pb_stats['q1']:.2f}",
                    f"{industry_pb_stats['q3']:.2f}",
                    f"{industry_pb_stats['std']:.2f}",
                    f"{((current_pb - industry_pb_stats['mean']) / industry_pb_stats['mean'] * 100):+.1f}%" if industry_pb_stats['mean'] != 0 else 'N/A',
                ],
            ]
            add_table_data(document, headers, data)

            add_paragraph(document, '')
            add_paragraph(document, f'同行公司数量: {len(pv)}家（过滤后PE有效: {len(pv_pe)}家, PB有效: {len(pv_pb)}家）')

            # 行业估值区间判断
            if current_pe > 0 and industry_pe_stats['mean'] > 0:
                if current_pe < industry_pe_stats['q1']:
                    pe_comment = '低于行业Q1，相对低估'
                elif current_pe > industry_pe_stats['q3']:
                    pe_comment = '高于行业Q3，相对高估'
                else:
                    pe_comment = '处于行业Q1-Q3区间，估值合理'
                add_paragraph(document, f'PE估值位置: {pe_comment}')

            if current_pb > 0 and industry_pb_stats['mean'] > 0:
                if current_pb < industry_pb_stats['q1']:
                    pb_comment = '低于行业Q1，相对低估'
                elif current_pb > industry_pb_stats['q3']:
                    pb_comment = '高于行业Q3，相对高估'
                else:
                    pb_comment = '处于行业Q1-Q3区间，估值合理'
                add_paragraph(document, f'PB估值位置: {pb_comment}')

        # 同行明细表
        if 'name' in pv.columns and 'pe_ttm' in pv.columns:
            pv_sorted = pv.sort_values('pe_ttm', ascending=True)
            has_level = 'peer_level' in pv_sorted.columns
            peer_headers = ['公司名称', '股票代码', '行业级别', 'PE(TTM)', 'PB', 'PS(TTM)', '市值(万元)'] if has_level else ['公司名称', '股票代码', 'PE(TTM)', 'PB', 'PS(TTM)', '市值(万元)']
            peer_rows = []
            for _, row in pv_sorted.head(30).iterrows():
                base = [
                    str(row.get('name', '')),
                    str(row.get('ts_code', '')),
                ]
                if has_level:
                    base.append(str(row.get('peer_level', '') or ''))
                base.extend([
                    f"{row.get('pe_ttm', 0):.2f}" if pd.notna(row.get('pe_ttm')) else '',
                    f"{row.get('pb', 0):.2f}" if pd.notna(row.get('pb')) else '',
                    f"{row.get('ps_ttm', 0):.2f}" if pd.notna(row.get('ps_ttm')) else '',
                    f"{row.get('total_mv', 0):.0f}" if pd.notna(row.get('total_mv')) else '',
                ])
                peer_rows.append(base)
            if peer_rows:
                add_paragraph(document, '')
                add_paragraph(document, '同行估值明细（按PE排序）:')
                add_table_data(document, peer_headers, peer_rows)
    else:
        add_paragraph(document, '同行估值数据不足，跳过同行对比分析。')

    # ------------------------------------------------------------------
    # 5.3  SW行业指数PE比较
    # ------------------------------------------------------------------
    add_title(document, '5.3 SW行业指数PE比较', level=2)
    add_paragraph(document, (
        '通过tushare sw_daily接口获取申万三级行业指数的PE数据，'
        '将标的个股PE与行业指数PE进行对比。'
    ))

    sw_index_pe = None
    sw_index_pb = None
    l3_code = sw_industry.get('l3_code', '')
    l3_name = sw_industry.get('l3_name', '')

    if l3_code and pro is not None:
        try:
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=30)).strftime('%Y%m%d')
            df_sw = pro.sw_daily(
                ts_code=l3_code, start_date=start_date, end_date=end_date
            )
            if df_sw is not None and not df_sw.empty:
                latest_sw = df_sw.iloc[-1]
                sw_index_pe = latest_sw.get('pe', None)
                sw_index_pb = latest_sw.get('pb', None)
        except Exception as e:
            print(f"  获取SW行业指数失败: {e}")

    if sw_index_pe is not None:
        current_pe_val = daily_basic.get('pe_ttm', 0) or 0
        pe_deviation = (current_pe_val - sw_index_pe) / sw_index_pe * 100 if sw_index_pe else 0

        headers = ['指标', f'{stock_name}', f'申万行业({l3_name})', '偏离度']
        data = [
            [
                'PE(TTM)',
                f"{current_pe_val:.2f}",
                f"{sw_index_pe:.2f}",
                f"{pe_deviation:+.1f}%",
            ],
        ]
        if sw_index_pb is not None:
            current_pb_val = daily_basic.get('pb', 0) or 0
            pb_deviation = (current_pb_val - sw_index_pb) / sw_index_pb * 100 if sw_index_pb else 0
            data.append([
                'PB',
                f"{current_pb_val:.2f}",
                f"{sw_index_pb:.2f}",
                f"{pb_deviation:+.1f}%",
            ])
        add_table_data(document, headers, data)

        add_paragraph(document, '')
        if pe_deviation > 20:
            add_paragraph(document, f'标的PE相对行业指数溢价{pe_deviation:.1f}%，估值偏高。')
        elif pe_deviation < -20:
            add_paragraph(document, f'标的PE相对行业指数折价{abs(pe_deviation):.1f}%，估值偏低。')
        else:
            add_paragraph(document, f'标的PE相对行业指数偏离度{pe_deviation:+.1f}%，估值基本合理。')
    else:
        add_paragraph(document, f'申万行业指数{l3_name}({l3_code})的PE数据暂不可获取。')

    # ------------------------------------------------------------------
    # 5.4  估值水平判定
    # ------------------------------------------------------------------
    add_title(document, '5.4 估值水平判定', level=2)
    add_paragraph(document, (
        '基于PE历史百分位进行估值水平判定。判定标准：\n'
        '- 百分位 < 20%：低估\n'
        '- 百分位 20%-40%：偏低估\n'
        '- 百分位 40%-60%：合理\n'
        '- 百分位 60%-80%：偏高估\n'
        '- 百分位 > 80%：高估'
    ))

    valuation_level = '无法判定'
    if pe_percentile is not None:
        if pe_percentile < 20:
            valuation_level = '低估'
        elif pe_percentile < 40:
            valuation_level = '偏低估'
        elif pe_percentile < 60:
            valuation_level = '合理'
        elif pe_percentile < 80:
            valuation_level = '偏高估'
        else:
            valuation_level = '高估'

        add_paragraph(document, f'PE历史百分位: {pe_percentile:.1f}%')
        if pb_percentile is not None:
            add_paragraph(document, f'PB历史百分位: {pb_percentile:.1f}%')
        add_paragraph(document, f'综合估值水平判定: {valuation_level}')

        add_paragraph(document, '')
        if valuation_level == '低估':
            add_paragraph(document, '当前估值处于历史低位，安全边际较高，具备较好的投资价值。')
        elif valuation_level == '偏低估':
            add_paragraph(document, '当前估值处于历史中低水平，具备一定安全边际。')
        elif valuation_level == '合理':
            add_paragraph(document, '当前估值处于历史中间水平，估值相对合理。')
        elif valuation_level == '偏高估':
            add_paragraph(document, '当前估值处于历史中高水平，需注意估值回调风险。')
        elif valuation_level == '高估':
            add_paragraph(document, '当前估值处于历史高位，估值回调风险较大，需谨慎对待。')
    else:
        add_paragraph(document, '由于历史PE数据不足，无法进行估值水平判定。')

    # 计算相对估值目标价
    target_price = 0
    current_pe_val = float(daily_basic.get('pe_ttm', 0) or 0)
    current_pb_val = float(daily_basic.get('pb', 0) or 0)
    current_price_val = float(daily_basic.get('close', 0) or 0)
    target_prices = []
    if pe_stats.get('p50') and current_pe_val > 0 and current_price_val > 0:
        target_prices.append(current_price_val * pe_stats['p50'] / current_pe_val)
    if pb_stats.get('p50') and current_pb_val > 0 and current_price_val > 0:
        target_prices.append(current_price_val * pb_stats['p50'] / current_pb_val)
    if not target_prices and industry_pe_stats.get('median') and current_pe_val > 0 and current_price_val > 0:
        target_prices.append(current_price_val * industry_pe_stats['median'] / current_pe_val)
    if target_prices:
        target_price = float(np.mean(target_prices))

    add_section_break(document)

    # ==================== 保存结果 ====================
    context['results']['relative_valuation'] = {
        'pe_percentile': pe_percentile,
        'pb_percentile': pb_percentile,
        'valuation_level': valuation_level,
        'pe_stats': pe_stats,
        'pb_stats': pb_stats,
        'industry_pe_stats': industry_pe_stats,
        'industry_pb_stats': industry_pb_stats,
        'target_price': target_price,
    }

    return context
