#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第三章：行业分析
SW行业分类表、同行PE/PB/PS对比表、行业指数PE百分位分析、PE趋势对比图
"""

import os
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

import module_utils as utils


def generate_chapter(context):
    """第三章：行业分析"""
    document = context['document']
    stock_code = context['stock_code']
    stock_name = context['stock_name']
    font_prop = context['font_prop']
    sw_industry = context.get('sw_industry', {})
    peer_valuation = context.get('peer_valuation', pd.DataFrame())
    historical_valuation = context.get('historical_valuation', pd.DataFrame())
    daily_basic = context.get('daily_basic', {})
    IMAGES_DIR = context['IMAGES_DIR']

    utils.add_title(document, '第三章 行业分析', level=1)

    # ==================== 3.1 SW行业分类表 ====================
    utils.add_title(document, '3.1 申万行业分类', level=2)

    l1_name = sw_industry.get('l1_name', '-')
    l2_name = sw_industry.get('l2_name', '-')
    l3_name = sw_industry.get('l3_name', '-')
    l1_code = sw_industry.get('l1_code', '-')
    l2_code = sw_industry.get('l2_code', '-')
    l3_code = sw_industry.get('l3_code', '-')

    sw_headers = ['级别', '行业名称', '行业代码']
    sw_data = [
        ['一级行业', l1_name, l1_code],
        ['二级行业', l2_name, l2_code],
        ['三级行业', l3_name, l3_code],
    ]
    utils.add_table_data(document, sw_headers, sw_data)

    utils.add_paragraph(
        document,
        f'{stock_name}所属申万一级行业为"{l1_name}"，二级行业为"{l2_name}"，'
        f'三级行业为"{l3_name}"（代码：{l3_code}）。'
    )

    # ==================== 3.2 同行PE/PB/PS对比表 ====================
    utils.add_title(document, '3.2 同行估值对比', level=2)

    has_peer_level = 'peer_level' in peer_valuation.columns if not peer_valuation.empty else False

    if has_peer_level:
        level_summary = peer_valuation['peer_level'].value_counts()
        level_desc = '、'.join([f'{k} {v}家' for k, v in level_summary.items()])
        utils.add_paragraph(document, f'同行来源：{level_desc}')
        utils.add_paragraph(document, (
            '说明：L3表示申万三级行业直接同行，L3兄弟表示同一二级行业下的其他三级子行业，'
            'L2表示二级行业全量成员。不同来源的同行业可比性存在差异，仅供参考。'
        ))

    if not peer_valuation.empty:
        # Prepare peer comparison table
        peer_headers = ['公司', '行业级别', 'PE(TTM)', 'PB', 'PS(TTM)', '总市值(万元)']

        # Add current stock as the first row
        current_pe = daily_basic.get('pe_ttm', '')
        current_pb = daily_basic.get('pb', '')
        current_ps = daily_basic.get('ps_ttm', '')
        current_mv = daily_basic.get('total_mv', 0)

        peer_data = []
        peer_data.append([
            f'{stock_name}(标的)',
            f'L3({sw_industry.get("l3_name", "")})' if sw_industry.get('l3_name') else '-',
            f'{current_pe:.2f}' if isinstance(current_pe, (int, float)) and current_pe else str(current_pe),
            f'{current_pb:.2f}' if isinstance(current_pb, (int, float)) and current_pb else str(current_pb),
            f'{current_ps:.2f}' if isinstance(current_ps, (int, float)) and current_ps else str(current_ps),
            f'{current_mv:,.2f}' if isinstance(current_mv, (int, float)) and current_mv else str(current_mv),
        ])

        # Sort peers by total_mv descending
        sorted_peers = peer_valuation.sort_values('total_mv', ascending=False) if 'total_mv' in peer_valuation.columns else peer_valuation

        for _, row in sorted_peers.iterrows():
            name = row.get('name', row.get('ts_code', ''))
            pe = row.get('pe_ttm', None)
            pb = row.get('pb', None)
            ps = row.get('ps_ttm', None)
            mv = row.get('total_mv', None)
            level = row.get('peer_level', '')

            peer_data.append([
                str(name),
                str(level) if level else '-',
                f'{pe:.2f}' if pd.notna(pe) and isinstance(pe, (int, float)) else '-',
                f'{pb:.2f}' if pd.notna(pb) and isinstance(pb, (int, float)) else '-',
                f'{ps:.2f}' if pd.notna(ps) and isinstance(ps, (int, float)) else '-',
                f'{mv:,.2f}' if pd.notna(mv) and isinstance(mv, (int, float)) else '-',
            ])

        utils.add_table_data(document, peer_headers, peer_data)

        # Compute peer statistics
        if 'pe_ttm' in peer_valuation.columns:
            valid_pe = peer_valuation['pe_ttm'].dropna()
            valid_pe = valid_pe[valid_pe > 0]
            if not valid_pe.empty:
                pe_median = valid_pe.median()
                pe_mean = valid_pe.mean()
                pe_min = valid_pe.min()
                pe_max = valid_pe.max()

                stat_headers = ['统计指标', 'PE(TTM)', 'PB', 'PS(TTM)']
                stat_data = []

                pe_vals = peer_valuation['pe_ttm'].dropna()
                pe_vals = pe_vals[pe_vals > 0]
                pb_vals = peer_valuation['pb'].dropna()
                pb_vals = pb_vals[pb_vals > 0]
                ps_vals = peer_valuation['ps_ttm'].dropna()
                ps_vals = ps_vals[ps_vals > 0]

                for stat_name, stat_func in [('平均值', np.mean), ('中位数', np.median), ('最小值', np.min), ('最大值', np.max)]:
                    row = [stat_name]
                    for vals in [pe_vals, pb_vals, ps_vals]:
                        if not vals.empty:
                            row.append(f'{stat_func(vals):.2f}')
                        else:
                            row.append('-')
                    stat_data.append(row)

                utils.add_paragraph(document, '')
                utils.add_paragraph(document, '同行估值统计：', bold=True)
                utils.add_table_data(document, stat_headers, stat_data)
        else:
            utils.add_paragraph(document, '同行数据中无PE(TTM)字段。')
    else:
        utils.add_paragraph(document, '暂无同行估值数据。')

    # ==================== 3.3 行业指数PE百分位分析 ====================
    utils.add_title(document, '3.3 历史PE百分位分析', level=2)

    if not historical_valuation.empty and 'pe_ttm' in historical_valuation.columns:
        pe_series = historical_valuation['pe_ttm'].dropna()
        pe_series = pe_series[pe_series > 0]  # Exclude negative PE

        if not pe_series.empty:
            # Compute percentiles
            current_pe_val = daily_basic.get('pe_ttm', None)

            pct_10 = pe_series.quantile(0.10)
            pct_25 = pe_series.quantile(0.25)
            pct_50 = pe_series.quantile(0.50)
            pct_75 = pe_series.quantile(0.75)
            pct_90 = pe_series.quantile(0.90)
            pct_mean = pe_series.mean()

            # Current PE percentile
            if current_pe_val and isinstance(current_pe_val, (int, float)) and current_pe_val > 0:
                current_pct = (pe_series < current_pe_val).mean() * 100
            else:
                current_pct = None

            pct_headers = ['百分位', 'PE(TTM)']
            pct_data = [
                ['10%分位', f'{pct_10:.2f}'],
                ['25%分位', f'{pct_25:.2f}'],
                ['50%分位(中位数)', f'{pct_50:.2f}'],
                ['75%分位', f'{pct_75:.2f}'],
                ['90%分位', f'{pct_90:.2f}'],
                ['平均值', f'{pct_mean:.2f}'],
                ['当前PE', f'{current_pe_val:.2f}' if isinstance(current_pe_val, (int, float)) and current_pe_val else '-'],
                ['当前百分位', f'{current_pct:.1f}%' if current_pct is not None else '-'],
            ]
            utils.add_table_data(document, pct_headers, pct_data)

            # Interpretation
            if current_pct is not None:
                if current_pct < 25:
                    verdict = '处于历史低位，估值较低'
                elif current_pct < 50:
                    verdict = '处于历史偏低位置'
                elif current_pct < 75:
                    verdict = '处于历史中位水平'
                else:
                    verdict = '处于历史高位，估值偏高'

                utils.add_paragraph(
                    document,
                    f'当前PE(TTM)为{current_pe_val:.2f}，在近3年数据中处于{current_pct:.1f}%百分位，'
                    f'{verdict}。'
                )
        else:
            utils.add_paragraph(document, '历史PE数据不足，无法计算百分位。')
    else:
        utils.add_paragraph(document, '暂无历史估值数据。')

    # ==================== 3.4 PE趋势对比图 ====================
    utils.add_title(document, '3.4 PE趋势与百分位图', level=2)

    if not historical_valuation.empty and 'pe_ttm' in historical_valuation.columns:
        pe_series = historical_valuation.set_index('trade_date')['pe_ttm'] if 'trade_date' in historical_valuation.columns else historical_valuation['pe_ttm']
        pe_series = pe_series.dropna()
        pe_series = pe_series[pe_series > 0]

        if not pe_series.empty and len(pe_series) > 10:
            chart_path = os.path.join(IMAGES_DIR, f'{stock_code}_pe_percentile.png')
            result = utils.generate_pe_percentile_chart(
                pe_series, chart_path, stock_name=stock_name, window=min(250, len(pe_series) // 2)
            )
            if result:
                utils.add_image(document, chart_path, width=utils.Inches(5.5))
                utils.add_paragraph(document, f'图3-1 {stock_name} PE(TTM)历史走势与百分位分析')
        else:
            utils.add_paragraph(document, '历史PE数据不足，无法生成趋势图。')
    else:
        utils.add_paragraph(document, '暂无历史估值数据，无法生成PE趋势图。')

    utils.add_section_break(document)

    return context
