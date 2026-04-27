#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第一章：项目概述
封面、目录、公司基本信息、当前行情概览、近3个月股价走势图
"""

import os
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

import module_utils as utils


def generate_chapter(context):
    """第一章：项目概述"""
    document = context['document']
    stock_code = context['stock_code']
    stock_name = context['stock_name']
    report_date = context['report_date']
    font_prop = context['font_prop']
    company_info = context.get('company_info', {})
    daily_basic = context.get('daily_basic', {})
    price_data = context.get('price_data', pd.DataFrame())
    IMAGES_DIR = context['IMAGES_DIR']

    # ==================== 封面 ====================
    utils.add_title(document, '上市公司估值分析报告', level=0)
    utils.add_paragraph(document, '')

    # 基本信息表
    industry_name = company_info.get('industry', '')
    market_name = company_info.get('market', '')
    pe_ttm = daily_basic.get('pe_ttm', '')
    pb = daily_basic.get('pb', '')
    total_mv = daily_basic.get('total_mv', 0)

    cover_headers = ['项目', '内容']
    cover_data = [
        ['公司名称', stock_name],
        ['股票代码', stock_code],
        ['报告日期', report_date],
        ['所属行业', industry_name],
        ['上市板块', market_name],
        ['PE(TTM)', f'{pe_ttm:.2f}' if isinstance(pe_ttm, (int, float)) and pe_ttm else str(pe_ttm)],
        ['PB', f'{pb:.2f}' if isinstance(pb, (int, float)) and pb else str(pb)],
        ['总市值(万元)', f'{total_mv:,.2f}' if isinstance(total_mv, (int, float)) and total_mv else str(total_mv)],
    ]
    utils.add_table_data(document, cover_headers, cover_data)

    utils.add_section_break(document)

    # ==================== 目录 ====================
    utils.add_title(document, '目  录', level=1)
    chapter_titles = [
        '第一章  项目概述',
        '第二章  财务分析',
        '第三章  行业分析',
        '第四章  DCF绝对估值',
        '第五章  相对估值',
        '第六章  敏感性分析',
        '第七章  蒙特卡洛模拟',
        '第八章  情景分析',
        '第九章  压力测试',
        '第十章  VaR风险度量',
        '第十一章  综合评估',
    ]
    for title in chapter_titles:
        utils.add_paragraph(document, title)

    utils.add_section_break(document)

    # ==================== 1.1 公司基本信息表 ====================
    utils.add_title(document, '1.1 公司基本信息', level=2)

    list_date = company_info.get('list_date', '')
    if isinstance(list_date, str) and len(list_date) == 8:
        list_date = f'{list_date[:4]}-{list_date[4:6]}-{list_date[6:8]}'

    info_headers = ['项目', '内容']
    info_data = [
        ['公司名称', company_info.get('name', stock_name)],
        ['股票代码', stock_code],
        ['英文名称', company_info.get('enname', '')],
        ['行业', industry_name],
        ['上市板块', market_name],
        ['上市日期', list_date],
        ['董事长', company_info.get('chairman', '')],
        ['省份', company_info.get('province', '')],
        ['城市', company_info.get('city', '')],
        ['公司网址', company_info.get('website', '')],
    ]
    utils.add_table_data(document, info_headers, info_data)

    # ==================== 1.2 当前行情概览表 ====================
    utils.add_title(document, '1.2 当前行情概览', level=2)

    close_price = daily_basic.get('close', '')
    ps_ttm = daily_basic.get('ps_ttm', '')
    total_share = daily_basic.get('total_share', 0)
    trade_date = daily_basic.get('trade_date', '')
    if isinstance(trade_date, str) and len(trade_date) == 8:
        trade_date = f'{trade_date[:4]}-{trade_date[4:6]}-{trade_date[6:8]}'

    # 计算流通股（总股本 - 限售股本，若无直接字段则使用总股本）
    float_share = daily_basic.get('float_share', total_share)

    market_headers = ['指标', '数值']
    market_data = [
        ['最新股价(元)', f'{close_price:.2f}' if isinstance(close_price, (int, float)) and close_price else str(close_price)],
        ['PE(TTM)', f'{pe_ttm:.2f}' if isinstance(pe_ttm, (int, float)) and pe_ttm else str(pe_ttm)],
        ['PB', f'{pb:.2f}' if isinstance(pb, (int, float)) and pb else str(pb)],
        ['PS(TTM)', f'{ps_ttm:.2f}' if isinstance(ps_ttm, (int, float)) and ps_ttm else str(ps_ttm)],
        ['总市值(万元)', f'{total_mv:,.2f}' if isinstance(total_mv, (int, float)) and total_mv else str(total_mv)],
        ['总股本(万股)', f'{total_share:,.2f}' if isinstance(total_share, (int, float)) and total_share else str(total_share)],
        ['流通股本(万股)', f'{float_share:,.2f}' if isinstance(float_share, (int, float)) and float_share else str(float_share)],
        ['数据日期', trade_date],
    ]
    utils.add_table_data(document, market_headers, market_data)

    # ==================== 1.3 近3个月股价走势图 ====================
    utils.add_title(document, '1.3 近3个月股价走势', level=2)

    if not price_data.empty:
        chart_path = os.path.join(IMAGES_DIR, f'{stock_code}_price_trend.png')
        result = utils.generate_price_trend_chart(price_data, chart_path, stock_name=stock_name)
        if result:
            utils.add_image(document, chart_path, width=utils.Inches(5.5))
            utils.add_paragraph(document, f'图1-1 {stock_name}近3个月股价走势与均线分析')
    else:
        utils.add_paragraph(document, '暂无近3个月股价数据。')

    utils.add_section_break(document)

    return context
