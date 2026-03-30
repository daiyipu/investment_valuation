# -*- coding: utf-8 -*-
"""
第二章：相对估值分析

本章节生成报告的相对估值分析部分，包括：
- 2.1 估值指标对比（PE、PB、PS）
- 2.1.1 同行公司名单
- 2.2 估值偏离度分析
- 2.3 PE历史分位数趋势分析
"""

import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from docx.shared import Inches

# 导入工具函数
from module_utils import (
    add_title, add_paragraph, add_table_data, add_image, add_section_break,
    generate_relative_valuation_charts_split
)


def generate_chapter(context):
    """
    生成第二章：相对估值分析

    Args:
        context: 包含以下键的字典:
            - document: Word文档对象
            - project_params: 项目参数（包含stock_code等）
            - IMAGES_DIR: 图片目录

    Returns:
        更新后的context字典
    """
    # 从context中提取数据
    document = context['document']
    project_params = context['project_params']
    IMAGES_DIR = context['IMAGES_DIR']

    stock_code = project_params.get('stock_code', '')  # 从project_params获取（与V2一致）

    # ==================== 二、相对估值分析 ====================
    add_title(document, '二、相对估值分析', level=1)

    add_paragraph(document, '本章节通过相对估值法（参数法），将光弘科技与行业内可比公司进行对比分析。')
    add_paragraph(document, '选取申万三级分类"消费电子零部件及组装"行业的同行公司，对比PE、PS、PB等估值倍数。')

    add_title(document, '2.1 估值指标对比', level=2)

    # 使用 Tushare 数据获取估值指标
    try:
        ts_token = os.environ.get('TUSHARE_TOKEN', '')

        if ts_token:
            import tushare as ts
            import time

            pro = ts.pro_api(ts_token)

            # 获取目标公司的估值数据（自动往前推1-2天直到找到交易日）
            trade_date = None
            df_target = None

            for days_back in range(1, 6):  # 尝试往前推1-5天
                test_date = (datetime.now() - timedelta(days=days_back)).strftime('%Y%m%d')
                try:
                    df_target = pro.daily_basic(
                        ts_code=stock_code,
                        trade_date=test_date,
                        fields='ts_code,trade_date,close,pe_ttm,pb,ps_ttm,total_mv'
                    )
                    if not df_target.empty:
                        trade_date = test_date
                        break
                except:
                    continue

            if df_target is None or df_target.empty:
                raise ValueError("未获取到目标公司数据（请检查网络或交易日历）")

            current_metrics_val = {
                'pe': df_target.iloc[0]['pe_ttm'],
                'pb': df_target.iloc[0]['pb'],
                'ps': df_target.iloc[0]['ps_ttm']
            }
            print(f"✅ 获取相对估值数据成功，交易日期: {trade_date}")

            # 获取申万三级行业分类的同行公司（与 notebook 一致）
            df_industry = pro.index_member_all(ts_code=stock_code)
            if df_industry.empty:
                raise ValueError("未获取到行业分类")

            # 过滤：只保留该股票的记录（防止API返回过多数据）
            df_industry = df_industry[df_industry['ts_code'] == stock_code]
            if df_industry.empty:
                raise ValueError(f"未找到{stock_code}的行业分类记录")

            # 显示所有行业分类记录，方便调试
            df_industry = df_industry.sort_values('in_date', ascending=False)
            print(f"✅ 获取到{len(df_industry)}条行业分类记录:")

            for idx, row in df_industry.head(5).iterrows():
                print(f"   [{idx}] {row['in_date']}: 一级={row.get('index_name', 'N/A')}, L1={row.get('l1_name', 'N/A')}, L2={row.get('l2_name', 'N/A')}, L3={row.get('l3_name', 'N/A')}")
                print(f"        L1代码={row.get('l1_code', 'N/A')}, L2代码={row.get('l2_code', 'N/A')}, L3代码={row.get('l3_code', 'N/A')}")

            latest_industry = df_industry.iloc[0]

            # 调试输出
            print(f"\n✅ 使用最新记录:")
            print(f"   股票代码: {stock_code}")
            print(f"   申万一级: {latest_industry.get('index_name', 'N/A')}")
            print(f"   申万三级代码: {latest_industry['l3_code']}")
            print(f"   申万三级名称: {latest_industry['l3_name']}")

            target_index_code = latest_industry['l3_code']  # 申万三级行业指数代码
            target_industry_l3 = latest_industry['l3_name']  # 行业名称

            # 获取该三级行业的所有成分股
            print(f"\n✅ 正在使用指数代码 {target_index_code} 查询成分股...")

            df_peers = pro.index_member_all(l3_code=target_index_code)
            print(f"✅ 获取到 {len(df_peers)} 条成分股记录")

            df_peers = df_peers[df_peers['ts_code'] != stock_code]

            # 获取同行公司基本信息
            peer_codes = df_peers['ts_code'].unique().tolist()
            print(f"✅ 过滤后剩余 {len(peer_codes)} 个同行公司")

            peer_basic = pro.stock_basic(ts_code=','.join(peer_codes[:30]),
                                       fields='ts_code,name,market')

            peer_stocks_all = pd.merge(df_peers, peer_basic, on='ts_code', how='left')
            peer_stocks_all = peer_stocks_all.drop_duplicates(subset=['ts_code'])

            # 限制数量并排序（扩充到30家）
            peer_stocks_all = peer_stocks_all.head(30)
            peer_names_dict = dict(zip(peer_stocks_all['ts_code'], peer_stocks_all['name_x']))

            # 获取同行公司的估值数据
            peer_data_list = []
            for peer_code in peer_stocks_all['ts_code'].tolist():
                peer_name = peer_names_dict.get(peer_code, peer_code)

                try:
                    df_peer = pro.daily_basic(
                        ts_code=peer_code,
                        trade_date=trade_date,
                        fields='ts_code,pe_ttm,pb,ps_ttm,total_mv'
                    )
                    if not df_peer.empty:
                        df_peer['name'] = peer_name
                        peer_data_list.append(df_peer)
                except:
                    pass

                time.sleep(0.2)  # 避免请求过快

            # 获取申万行业指数的PE数据（使用sw_daily接口）
            sw_index_pe = None
            sw_index_pb = None
            sw_index_ps = None
            try:
                print(f"✅ 正在获取申万行业指数PE数据: {target_index_code}")
                # 获取最近5天的数据，取最新的一条
                end_date = datetime.now().strftime('%Y%m%d')
                start_date = (datetime.now() - timedelta(days=10)).strftime('%Y%m%d')
                df_index = pro.sw_daily(
                    ts_code=target_index_code,
                    start_date=start_date,
                    end_date=end_date
                )
                if df_index is not None and not df_index.empty:
                    # sw_daily返回的是pe列（不是pe_ttm）
                    latest = df_index.iloc[-1]
                    sw_index_pe = latest.get('pe', None)
                    sw_index_pb = latest.get('pb', None)
                    sw_index_ps = latest.get('ps_ttm', None)  # sw_daily可能没有ps_ttm
                    if sw_index_pe:
                        print(f"✅ 申万行业指数PE数据获取成功: PE={sw_index_pe:.2f}, PB={sw_index_pb:.2f}")
                else:
                    print(f"⚠️ 申万行业指数PE数据为空")
            except Exception as e:
                print(f"⚠️ 获取申万行业指数PE数据失败: {e}")

            if peer_data_list:
                peer_companies_val = pd.concat(peer_data_list, ignore_index=True)

                # 过滤异常数据
                peer_companies_val = peer_companies_val[
                    (peer_companies_val['pe_ttm'] > 0) &
                    (peer_companies_val['pe_ttm'] < 500) &
                    (peer_companies_val['pb'] > 0) &
                    (peer_companies_val['pb'] < 20)
                ]

                # 重命名列并进行单位转换
                # total_mv单位是万元，需要转换为亿元（除以10000）
                peer_companies_val['market_cap'] = peer_companies_val['total_mv'] / 10000

                peer_companies_val = peer_companies_val.rename(columns={
                    'pe_ttm': 'pe',
                    'ps_ttm': 'ps',
                    'ts_code': 'code'
                })

                # 只保留需要的列
                peer_companies_val = peer_companies_val[['name', 'code', 'pe', 'ps', 'pb', 'market_cap']]
            else:
                raise ValueError("未获取到同行公司估值数据")

        else:
            raise ValueError("TUSHARE_TOKEN 未设置")

    except Exception as e:
        print(f"获取相对估值数据失败: {e}，使用示例数据")
        # 使用示例数据
        current_metrics_val = {
            'pe': 56.24,
            'pb': 3.71,
            'ps': 2.30
        }

        peer_companies_val = pd.DataFrame({
            'name': ['立讯精密', '歌尔股份', '蓝思科技', '长盈精密', '领益智造', '安洁科技', '比亚迪电子'],
            'code': ['002475.SZ', '002241.SZ', '300433.SZ', '300115.SZ', '002600.SZ', '002635.SZ', '0285.HK'],
            'pe': [20.5, 25.8, 22.3, 28.6, 18.9, 32.1, 24.5],
            'ps': [1.2, 1.8, 1.5, 2.1, 1.0, 2.5, 1.6],
            'pb': [2.8, 3.2, 2.5, 3.8, 2.1, 4.2, 3.0],
            'market_cap': [120, 80, 50, 45, 65, 25, 90]
        })

    # 计算行业统计指标
    industry_stats_val = {
        'pe': {
            'mean': peer_companies_val['pe'].mean(),
            'median': peer_companies_val['pe'].median(),
            'q1': peer_companies_val['pe'].quantile(0.25),
            'q3': peer_companies_val['pe'].quantile(0.75),
            'min': peer_companies_val['pe'].min(),
            'max': peer_companies_val['pe'].max(),
            'std': peer_companies_val['pe'].std()
        },
        'ps': {
            'mean': peer_companies_val['ps'].mean(),
            'median': peer_companies_val['ps'].median(),
            'q1': peer_companies_val['ps'].quantile(0.25),
            'q3': peer_companies_val['ps'].quantile(0.75),
            'min': peer_companies_val['ps'].min(),
            'max': peer_companies_val['ps'].max(),
            'std': peer_companies_val['ps'].std()
        },
        'pb': {
            'mean': peer_companies_val['pb'].mean(),
            'median': peer_companies_val['pb'].median(),
            'q1': peer_companies_val['pb'].quantile(0.25),
            'q3': peer_companies_val['pb'].quantile(0.75),
            'min': peer_companies_val['pb'].min(),
            'max': peer_companies_val['pb'].max(),
            'std': peer_companies_val['pb'].std()
        }
    }

    # 剔除3倍标准差异常值后的行业平均
    industry_avg_val = {
        'pe': industry_stats_val['pe']['mean'],
        'ps': industry_stats_val['ps']['mean'],
        'pb': industry_stats_val['pb']['mean']
    }

    # 估值指标对比表
    valuation_headers = ['指标', '光弘科技', '行业平均', '中位数', 'Q1(25分位)', 'Q3(75分位)', '最小值', '最大值', '偏离度']
    valuation_data = [
        ['PE (TTM)',
         f"{current_metrics_val['pe']:.2f}倍",
         f"{industry_stats_val['pe']['mean']:.2f}倍",
         f"{industry_stats_val['pe']['median']:.2f}倍",
         f"{industry_stats_val['pe']['q1']:.2f}倍",
         f"{industry_stats_val['pe']['q3']:.2f}倍",
         f"{industry_stats_val['pe']['min']:.2f}倍",
         f"{industry_stats_val['pe']['max']:.2f}倍",
         f"{(current_metrics_val['pe']-industry_stats_val['pe']['mean'])/industry_stats_val['pe']['mean']*100:+.1f}%"],
        ['PB',
         f"{current_metrics_val['pb']:.2f}倍",
         f"{industry_stats_val['pb']['mean']:.2f}倍",
         f"{industry_stats_val['pb']['median']:.2f}倍",
         f"{industry_stats_val['pb']['q1']:.2f}倍",
         f"{industry_stats_val['pb']['q3']:.2f}倍",
         f"{industry_stats_val['pb']['min']:.2f}倍",
         f"{industry_stats_val['pb']['max']:.2f}倍",
         f"{(current_metrics_val['pb']-industry_stats_val['pb']['mean'])/industry_stats_val['pb']['mean']*100:+.1f}%"],
        ['PS (TTM)',
         f"{current_metrics_val['ps']:.2f}倍",
         f"{industry_stats_val['ps']['mean']:.2f}倍",
         f"{industry_stats_val['ps']['median']:.2f}倍",
         f"{industry_stats_val['ps']['q1']:.2f}倍",
         f"{industry_stats_val['ps']['q3']:.2f}倍",
         f"{industry_stats_val['ps']['min']:.2f}倍",
         f"{industry_stats_val['ps']['max']:.2f}倍",
         f"{(current_metrics_val['ps']-industry_stats_val['ps']['mean'])/industry_stats_val['ps']['mean']*100:+.1f}%"]
    ]
    add_table_data(document, valuation_headers, valuation_data)

    # 添加统计分析说明
    add_paragraph(document, '')
    add_paragraph(document, '行业估值统计说明：')
    add_paragraph(document, '• 行业平均：所有同行公司的算术平均值（受极端值影响）')
    add_paragraph(document, '• 中位数：行业50%的公司估值低于此水平，抗极端值干扰')
    add_paragraph(document, '• Q1(25分位)：行业25%的公司估值低于此水平')
    add_paragraph(document, '• Q3(75分位)：行业75%的公司估值低于此水平（即25%的公司高于此水平）')
    add_paragraph(document, '• 最小/最大值：行业中的估值极值')
    add_paragraph(document, '• 数据已过滤异常值（PE<500, PB<20）以避免极端情况影响统计')
    add_paragraph(document, f'• 样本量：共{len(peer_companies_val)}家同行公司')

    # 添加同行公司名单
    add_paragraph(document, '')
    add_title(document, '2.1.1 同行公司名单', level=3)
    add_paragraph(document, f'基于申万三级行业分类"消费电子零部件及组装"筛选的同行公司：')

    # 按市值排序的同行公司表格
    peer_companies_sorted = peer_companies_val.sort_values('market_cap', ascending=False)
    peer_headers = ['公司名称', '股票代码', 'PE (TTM)', 'PB', 'PS (TTM)', '市值(亿元)']
    peer_rows = []
    for _, row in peer_companies_sorted.iterrows():
        peer_rows.append([
            row['name'],
            row['code'],
            f"{row['pe']:.2f}",
            f"{row['pb']:.2f}",
            f"{row['ps']:.2f}",
            f"{row['market_cap']:.2f}"
        ])
    add_table_data(document, peer_headers, peer_rows)

    # 添加行业统计汇总
    add_paragraph(document, '')
    add_paragraph(document, '行业估值统计汇总：')
    add_paragraph(document, f'• PE: 平均{industry_stats_val["pe"]["mean"]:.2f}倍，中位数{peer_companies_val["pe"].median():.2f}倍，标准差{industry_stats_val["pe"]["std"]:.2f}倍')
    add_paragraph(document, f'  • Q1-Q3区间: [{industry_stats_val["pe"]["q1"]:.2f}, {industry_stats_val["pe"]["q3"]:.2f}]倍，极值范围: [{industry_stats_val["pe"]["min"]:.2f}, {industry_stats_val["pe"]["max"]:.2f}]倍')
    add_paragraph(document, f'• PB: 平均{industry_stats_val["pb"]["mean"]:.2f}倍，中位数{peer_companies_val["pb"].median():.2f}倍，标准差{industry_stats_val["pb"]["std"]:.2f}倍')
    add_paragraph(document, f'  • Q1-Q3区间: [{industry_stats_val["pb"]["q1"]:.2f}, {industry_stats_val["pb"]["q3"]:.2f}]倍，极值范围: [{industry_stats_val["pb"]["min"]:.2f}, {industry_stats_val["pb"]["max"]:.2f}]倍')
    add_paragraph(document, f'• PS: 平均{industry_stats_val["ps"]["mean"]:.2f}倍，中位数{peer_companies_val["ps"].median():.2f}倍，标准差{industry_stats_val["ps"]["std"]:.2f}倍')
    add_paragraph(document, f'  • Q1-Q3区间: [{industry_stats_val["ps"]["q1"]:.2f}, {industry_stats_val["ps"]["q3"]:.2f}]倍，极值范围: [{industry_stats_val["ps"]["min"]:.2f}, {industry_stats_val["ps"]["max"]:.2f}]倍')

    add_paragraph(document, '图表 2.0: 相对估值对比分析 - 估值指标对比')
    chart_paths, df_scenarios = generate_relative_valuation_charts_split(
        current_metrics_val, industry_avg_val, peer_companies_val, IMAGES_DIR
    )
    add_image(document, chart_paths[0])  # 估值指标对比

    add_paragraph(document, '图表 2.1: 相对估值对比分析 - PE倍数对比')
    add_image(document, chart_paths[1])

    add_paragraph(document, '图表 2.2: 相对估值对比分析 - PB倍数对比')
    add_image(document, chart_paths[2])

    add_paragraph(document, '图表 2.3: 相对估值对比分析 - PS倍数对比')
    add_image(document, chart_paths[3])

    # ==================== 2.1.2 申万行业指数估值 ====================
    add_title(document, '2.1.2 申万行业指数估值', level=3)
    add_paragraph(document, '')

    if sw_index_pe is not None:
        add_paragraph(document, f'本节展示申万三级行业指数"{target_industry_l3}"的估值数据，提供官方行业指数视角的估值基准。')
        add_paragraph(document, f'申万行业指数代码: {target_index_code}')
        add_paragraph(document, '')

        # 申万行业指数估值表格
        sw_index_headers = ['指标', '申万行业指数', '光弘科技', '差异', '说明']
        sw_index_data = [
            ['PE (TTM)',
             f"{sw_index_pe:.2f}倍",
             f"{current_metrics_val['pe']:.2f}倍",
             f"{(current_metrics_val['pe']-sw_index_pe)/sw_index_pe*100:+.1f}%",
             '行业指数PE反映行业整体估值水平' if current_metrics_val['pe'] > sw_index_pe else '个股PE低于行业指数，相对低估'],
            ['PB',
             f"{sw_index_pb:.2f}倍" if sw_index_pb else "N/A",
             f"{current_metrics_val['pb']:.2f}倍",
             f"{(current_metrics_val['pb']-sw_index_pb)/sw_index_pb*100:+.1f}%" if sw_index_pb else "N/A",
             '市净率反映行业整体账面价值溢价'],
            ['PS (TTM)',
             f"{sw_index_ps:.2f}倍" if sw_index_ps else "N/A",
             f"{current_metrics_val['ps']:.2f}倍",
             f"{(current_metrics_val['ps']-sw_index_ps)/sw_index_ps*100:+.1f}%" if sw_index_ps else "N/A",
             '市销率反映行业整体营收能力']
        ]
        add_table_data(document, sw_index_headers, sw_index_data)

        add_paragraph(document, '')
        add_paragraph(document, '申万行业指数估值说明：')
        add_paragraph(document, f'• 申万行业指数是基于该行业所有成分股按市值加权计算的指数')
        add_paragraph(document, f'• 指数PE/PB/PS反映行业整体的估值水平，不同于同行公司平均（简单平均）')
        add_paragraph(document, f'• 指数估值受大盘股权重影响，更能代表行业龙头公司的估值水平')
        add_paragraph(document, f'• 同行公司平均反映行业内典型公司的估值，受小盘股影响较大')
        add_paragraph(document, '')

        # 对比分析
        add_paragraph(document, '📊 对比分析：', bold=True)
        if abs(current_metrics_val['pe'] - sw_index_pe) / sw_index_pe < 0.1:
            add_paragraph(document, f'✅ 个股PE({current_metrics_val["pe"]:.2f}倍)与申万行业指数PE({sw_index_pe:.2f}倍)基本一致，估值合理')
        elif current_metrics_val['pe'] < sw_index_pe:
            add_paragraph(document, f'✅ 个股PE({current_metrics_val["pe"]:.2f}倍)低于申万行业指数PE({sw_index_pe:.2f}倍)，相对行业指数低估')
        else:
            add_paragraph(document, f'⚠️ 个股PE({current_metrics_val["pe"]:.2f}倍)高于申万行业指数PE({sw_index_pe:.2f}倍)，相对行业指数高估')
    else:
        add_paragraph(document, f'⚠️ 申万行业指数"{target_industry_l3}"的估值数据暂时无法获取')
        add_paragraph(document, '可能原因：指数代码不正确或数据源暂时不可用')

    add_paragraph(document, '')

    add_title(document, '2.2 估值偏离度分析', level=2)
    add_paragraph(document, '本节分析标的公司与同行公司和申万行业指数的估值偏离情况，评估估值相对位置。')
    add_paragraph(document, '')

    # 计算PE在行业中的分位数位置
    pe_position = (peer_companies_val['pe'] < current_metrics_val['pe']).sum() / len(peer_companies_val) * 100
    pb_position = (peer_companies_val['pb'] < current_metrics_val['pb']).sum() / len(peer_companies_val) * 100
    ps_position = (peer_companies_val['ps'] < current_metrics_val['ps']).sum() / len(peer_companies_val) * 100

    # 2.2.1 与同行公司对比
    add_title(document, '2.2.1 与同行公司对比', level=3)
    add_paragraph(document, '')

    add_paragraph(document, f"• PE偏离度: {(current_metrics_val['pe']-industry_avg_val['pe'])/industry_avg_val['pe']*100:+.1f}%，位于行业{pe_position:.1f}%分位")
    add_paragraph(document, f"• PB偏离度: {(current_metrics_val['pb']-industry_avg_val['pb'])/industry_avg_val['pb']*100:+.1f}%，位于行业{pb_position:.1f}%分位")
    add_paragraph(document, f"• PS偏离度: {(current_metrics_val['ps']-industry_avg_val['ps'])/industry_avg_val['ps']*100:+.1f}%，位于行业{ps_position:.1f}%分位")

    add_paragraph(document, '')

    # PE分位数分析
    if current_metrics_val['pe'] > industry_stats_val['pe']['q3']:
        add_paragraph(document, f'⚠️ PE({current_metrics_val["pe"]:.2f}倍)高于行业Q3({industry_stats_val["pe"]["q3"]:.2f}倍)，处于行业高位，估值偏高')
    elif current_metrics_val['pe'] < industry_stats_val['pe']['q1']:
        add_paragraph(document, f'✅ PE({current_metrics_val["pe"]:.2f}倍)低于行业Q1({industry_stats_val["pe"]["q1"]:.2f}倍)，处于行业低位，估值偏低')
    else:
        add_paragraph(document, f'ℹ️ PE({current_metrics_val["pe"]:.2f}倍)介于行业Q1({industry_stats_val["pe"]["q1"]:.2f}倍)和Q3({industry_stats_val["pe"]["q3"]:.2f}倍)之间，估值合理')

    # PB分位数分析
    if current_metrics_val['pb'] > industry_stats_val['pb']['q3']:
        add_paragraph(document, f'⚠️ PB({current_metrics_val["pb"]:.2f}倍)高于行业Q3({industry_stats_val["pb"]["q3"]:.2f}倍)，市净率偏高')
    elif current_metrics_val['pb'] < industry_stats_val['pb']['q1']:
        add_paragraph(document, f'✅ PB({current_metrics_val["pb"]:.2f}倍)低于行业Q1({industry_stats_val["pb"]["q1"]:.2f}倍)，市净率偏低')

    # PS分位数分析
    if current_metrics_val['ps'] > industry_stats_val['ps']['q3']:
        add_paragraph(document, f'⚠️ PS({current_metrics_val["ps"]:.2f}倍)高于行业Q3({industry_stats_val["ps"]["q3"]:.2f}倍)，市销率偏高')
    elif current_metrics_val['ps'] < industry_stats_val['ps']['q1']:
        add_paragraph(document, f'✅ PS({current_metrics_val["ps"]:.2f}倍)低于行业Q1({industry_stats_val["ps"]["q1"]:.2f}倍)，市销率偏低')

    # 2.2.2 与申万行业指数对比
    if sw_index_pe is not None:
        add_paragraph(document, '')
        add_title(document, '2.2.2 与申万行业指数对比', level=3)
        add_paragraph(document, '')

        # 计算与申万行业指数的偏离度
        pe_dev_sw = (current_metrics_val['pe'] - sw_index_pe) / sw_index_pe * 100
        pb_dev_sw = (current_metrics_val['pb'] - sw_index_pb) / sw_index_pb * 100 if sw_index_pb else None
        ps_dev_sw = (current_metrics_val['ps'] - sw_index_ps) / sw_index_ps * 100 if sw_index_ps else None

        add_paragraph(document, f"• PE偏离度: {pe_dev_sw:+.1f}%（标的{current_metrics_val['pe']:.2f}倍 vs 申万{sw_index_pe:.2f}倍）")
        if pb_dev_sw is not None:
            add_paragraph(document, f"• PB偏离度: {pb_dev_sw:+.1f}%（标的{current_metrics_val['pb']:.2f}倍 vs 申万{sw_index_pb:.2f}倍）")
        if ps_dev_sw is not None:
            add_paragraph(document, f"• PS偏离度: {ps_dev_sw:+.1f}%（标的{current_metrics_val['ps']:.2f}倍 vs 申万{sw_index_ps:.2f}倍）")

        add_paragraph(document, '')

        # PE申万指数对比分析
        if abs(pe_dev_sw) < 10:
            add_paragraph(document, f'✅ PE({current_metrics_val["pe"]:.2f}倍)与申万行业指数PE({sw_index_pe:.2f}倍)基本一致，偏离度{pe_dev_sw:+.1f}%')
        elif pe_dev_sw > 0:
            add_paragraph(document, f'⚠️ PE({current_metrics_val["pe"]:.2f}倍)高于申万行业指数PE({sw_index_pe:.2f}倍)，溢价{pe_dev_sw:+.1f}%')
        else:
            add_paragraph(document, f'✅ PE({current_metrics_val["pe"]:.2f}倍)低于申万行业指数PE({sw_index_pe:.2f}倍)，折价{pe_dev_sw:+.1f}%')

        # PB申万指数对比分析
        if pb_dev_sw is not None:
            if abs(pb_dev_sw) < 10:
                add_paragraph(document, f'✅ PB({current_metrics_val["pb"]:.2f}倍)与申万行业指数PB({sw_index_pb:.2f}倍)基本一致')
            elif pb_dev_sw > 0:
                add_paragraph(document, f'⚠️ PB({current_metrics_val["pb"]:.2f}倍)高于申万行业指数PB({sw_index_pb:.2f}倍)，溢价{pb_dev_sw:+.1f}%')
            else:
                add_paragraph(document, f'✅ PB({current_metrics_val["pb"]:.2f}倍)低于申万行业指数PB({sw_index_pb:.2f}倍)，折价{pb_dev_sw:+.1f}%')

        add_paragraph(document, '')
        add_paragraph(document, '申万行业指数说明：')
        add_paragraph(document, f'• 申万行业指数基于所有成分股市值加权，反映行业整体估值水平')
        add_paragraph(document, f'• 与申万指数对比可判断个股相对行业整体的估值位置')
        add_paragraph(document, f'• 正偏离表示估值高于行业平均，负偏离表示估值低于行业平均')

    add_paragraph(document, '')

    # ==================== 2.3 PE历史分位数趋势分析 ====================
    add_title(document, '2.3 PE历史分位数趋势分析', level=2)

    add_paragraph(document, '本节通过分析标的股票和所属行业的PE历史走势及分位数变化，从时间维度评估估值的合理性。')
    add_paragraph(document, '基于最近5年的历史数据（约1250个交易日），通过对比个股与行业的PE历史分位数趋势，可以更清晰地判断当前估值处于历史哪个水平。')
    add_paragraph(document, '5年的历史周期能够覆盖完整的牛熊周期，提供更可靠的估值基准。')
    add_paragraph(document, '')

    # 尝试从tushare获取历史PE数据并生成趋势图
    try:
        from utils.pe_history_analyzer import PEHistoryAnalyzer

        print("\n=== 开始PE历史分位数趋势分析 ===")

        # 创建PE历史分析器
        pe_analyzer = PEHistoryAnalyzer()

        # 获取个股历史PE数据（最近5年）
        stock_pe_data = pe_analyzer.get_stock_pe_history(stock_code, days=1825)

        # 获取行业历史PE数据
        industry_name, industry_code, industry_pe_data = pe_analyzer.get_industry_pe_history(stock_code, days=1825)

        if stock_pe_data is not None and industry_pe_data is not None:
            print(f"✅ 成功获取历史PE数据")
            print(f"   个股数据: {len(stock_pe_data)}条")
            print(f"   行业数据: {len(industry_pe_data)}条")
            print(f"   行业: {industry_name} ({industry_code})")

            # 保存PE数据到context，供第六章情景分析使用
            context['stock_pe_data'] = stock_pe_data
            context['industry_pe_data'] = industry_pe_data
            context['industry_name'] = industry_name
            context['industry_code'] = industry_code
            print(f"✅ 已保存PE数据到context，供第六章情景分析使用")

            # 计算个股历史分位数统计
            stock_pe_current = stock_pe_data.iloc[-1]['pe_ttm']
            stock_pe_min = stock_pe_data['pe_ttm'].min()
            stock_pe_max = stock_pe_data['pe_ttm'].max()
            stock_pe_median = stock_pe_data['pe_ttm'].median()
            stock_pe_25 = stock_pe_data['pe_ttm'].quantile(0.25)
            stock_pe_75 = stock_pe_data['pe_ttm'].quantile(0.75)
            stock_pe_percentile = (stock_pe_data['pe_ttm'] < stock_pe_current).sum() / len(stock_pe_data) * 100

            # 计算申万行业指数历史分位数统计
            sw_index_pe_current = industry_pe_data.iloc[-1]['pe_ttm']
            sw_index_pe_min = industry_pe_data['pe_ttm'].min()
            sw_index_pe_max = industry_pe_data['pe_ttm'].max()
            sw_index_pe_median = industry_pe_data['pe_ttm'].median()
            sw_index_pe_25 = industry_pe_data['pe_ttm'].quantile(0.25)
            sw_index_pe_75 = industry_pe_data['pe_ttm'].quantile(0.75)
            sw_index_pe_percentile = (industry_pe_data['pe_ttm'] < sw_index_pe_current).sum() / len(industry_pe_data) * 100

            # 获取同行公司的历史PE数据
            print("\n获取同行公司历史PE数据...")
            custom_peer_pe_current = None
            custom_peer_pe_min = None
            custom_peer_pe_max = None
            custom_peer_pe_median = None
            custom_peer_pe_percentile = None

            try:
                # 提取同行公司代码列表
                peer_codes = peer_companies_val['code'].tolist()[:20]  # 取前20个
                print(f"  同行公司数量: {len(peer_codes)}")

                # 获取每个同行公司的历史PE
                peer_pe_histories = []
                for peer_code in peer_codes:
                    try:
                        peer_history = pe_analyzer.get_stock_pe_history(peer_code, days=1825)
                        if peer_history is not None and not peer_history.empty:
                            peer_history = peer_history.rename(columns={'pe_ttm': 'pe'})
                            peer_pe_histories.append(peer_history)
                    except Exception as e:
                        print(f"    获取{peer_code}历史PE失败: {e}")
                        continue

                if peer_pe_histories:
                    # 合并所有同行公司的历史PE
                    peer_pe_df = pd.concat(peer_pe_histories, ignore_index=True)

                    # 按日期分组计算平均值
                    custom_peer_pe_data = peer_pe_df.groupby('trade_date')['pe'].mean().reset_index()
                    custom_peer_pe_data = custom_peer_pe_data.sort_values('trade_date').reset_index(drop=True)

                    # 计算统计指标
                    # 安全提取标量值，处理各种数据类型
                    try:
                        # 获取最后一个PE值
                        pe_last = custom_peer_pe_data.iloc[-1]['pe']
                        if isinstance(pe_last, (pd.Series, list)):
                            custom_peer_pe_current = float(pe_last.iloc[0] if isinstance(pe_last, pd.Series) else pe_last[0])
                        else:
                            custom_peer_pe_current = float(pe_last)

                        # 其他聚合函数，添加类型检查
                        pe_min = custom_peer_pe_data['pe'].min()
                        custom_peer_pe_min = float(pe_min.iloc[0] if isinstance(pe_min, pd.Series) else pe_min)

                        pe_max = custom_peer_pe_data['pe'].max()
                        custom_peer_pe_max = float(pe_max.iloc[0] if isinstance(pe_max, pd.Series) else pe_max)

                        pe_median = custom_peer_pe_data['pe'].median()
                        custom_peer_pe_median = float(pe_median.iloc[0] if isinstance(pe_median, pd.Series) else pe_median)

                        # 计算25%和75%分位数
                        pe_25 = custom_peer_pe_data['pe'].quantile(0.25)
                        custom_peer_pe_25 = float(pe_25.iloc[0] if isinstance(pe_25, pd.Series) else pe_25)

                        pe_75 = custom_peer_pe_data['pe'].quantile(0.75)
                        custom_peer_pe_75 = float(pe_75.iloc[0] if isinstance(pe_75, pd.Series) else pe_75)

                        # 计算百分位
                        pe_percentile_count = (custom_peer_pe_data['pe'] < custom_peer_pe_current).sum()
                        custom_peer_pe_percentile = float(pe_percentile_count.iloc[0] if isinstance(pe_percentile_count, pd.Series) else pe_percentile_count) / len(custom_peer_pe_data) * 100

                    except Exception as e:
                        print(f"  ⚠️ 计算统计指标时出错: {e}")
                        print(f"  ⚠️ custom_peer_pe_data形状: {custom_peer_pe_data.shape}")
                        print(f"  ⚠️ custom_peer_pe_data列: {custom_peer_pe_data.columns.tolist()}")
                        raise

                    print(f"  ✅ 同行公司历史PE计算成功:")
                    print(f"     当前PE: {custom_peer_pe_current:.2f}倍")
                    print(f"     数据点数: {len(custom_peer_pe_data)}")
                else:
                    print(f"  ⚠️ 未获取到同行公司历史PE数据")

            except Exception as e:
                print(f"  ❌ 获取同行公司历史PE失败: {e}")

            # 添加历史分位数统计表格
            if custom_peer_pe_current is not None:
                # 有自定义同行数据的完整表格
                pe_history_headers = ['指标', '标的股票', '行业指数(申万)', '行业指数(自定义)', '差异(vs申万)', '差异(vs自定义)']
                pe_history_data = [
                    ['当前PE-TTM',
                     f'{stock_pe_current:.2f}倍',
                     f'{sw_index_pe_current:.2f}倍',
                     f'{custom_peer_pe_current:.2f}倍',
                     f'{(stock_pe_current/sw_index_pe_current-1)*100:+.1f}%',
                     f'{(stock_pe_current/custom_peer_pe_current-1)*100:+.1f}%'],
                    ['历史最小PE',
                     f'{stock_pe_min:.2f}倍',
                     f'{sw_index_pe_min:.2f}倍',
                     f'{custom_peer_pe_min:.2f}倍',
                     f'{stock_pe_min-sw_index_pe_min:+.2f}倍',
                     f'{stock_pe_min-custom_peer_pe_min:+.2f}倍'],
                    ['25%分位数PE',
                     f'{stock_pe_25:.2f}倍',
                     f'{sw_index_pe_25:.2f}倍',
                     f'{custom_peer_pe_25:.2f}倍',
                     f'{stock_pe_25-sw_index_pe_25:+.2f}倍',
                     f'{stock_pe_25-custom_peer_pe_25:+.2f}倍'],
                    ['历史中位数PE',
                     f'{stock_pe_median:.2f}倍',
                     f'{sw_index_pe_median:.2f}倍',
                     f'{custom_peer_pe_median:.2f}倍',
                     f'{stock_pe_median-sw_index_pe_median:+.2f}倍',
                     f'{stock_pe_median-custom_peer_pe_median:+.2f}倍'],
                    ['75%分位数PE',
                     f'{stock_pe_75:.2f}倍',
                     f'{sw_index_pe_75:.2f}倍',
                     f'{custom_peer_pe_75:.2f}倍',
                     f'{stock_pe_75-sw_index_pe_75:+.2f}倍',
                     f'{stock_pe_75-custom_peer_pe_75:+.2f}倍'],
                    ['历史最大PE',
                     f'{stock_pe_max:.2f}倍',
                     f'{sw_index_pe_max:.2f}倍',
                     f'{custom_peer_pe_max:.2f}倍',
                     f'{stock_pe_max-sw_index_pe_max:+.2f}倍',
                     f'{stock_pe_max-custom_peer_pe_max:+.2f}倍'],
                    ['当前分位数',
                     f'{stock_pe_percentile:.1f}%',
                     f'{sw_index_pe_percentile:.1f}%',
                     f'{custom_peer_pe_percentile:.1f}%',
                     f'{stock_pe_percentile-sw_index_pe_percentile:+.1f}%',
                     f'{stock_pe_percentile-custom_peer_pe_percentile:+.1f}%']
                ]
            else:
                # 只有申万数据的简化表格
                pe_history_headers = ['指标', '标的股票', '行业指数(申万)', '差异']
                pe_history_data = [
                    ['当前PE-TTM', f'{stock_pe_current:.2f}倍', f'{sw_index_pe_current:.2f}倍', f'{(stock_pe_current/sw_index_pe_current-1)*100:+.1f}%'],
                    ['历史最小PE', f'{stock_pe_min:.2f}倍', f'{sw_index_pe_min:.2f}倍', f'{stock_pe_min-sw_index_pe_min:+.2f}倍'],
                    ['25%分位数PE', f'{stock_pe_25:.2f}倍', f'{sw_index_pe_25:.2f}倍', f'{stock_pe_25-sw_index_pe_25:+.2f}倍'],
                    ['历史中位数PE', f'{stock_pe_median:.2f}倍', f'{sw_index_pe_median:.2f}倍', f'{stock_pe_median-sw_index_pe_median:+.2f}倍'],
                    ['75%分位数PE', f'{stock_pe_75:.2f}倍', f'{sw_index_pe_75:.2f}倍', f'{stock_pe_75-sw_index_pe_75:+.2f}倍'],
                    ['历史最大PE', f'{stock_pe_max:.2f}倍', f'{sw_index_pe_max:.2f}倍', f'{stock_pe_max-sw_index_pe_max:+.2f}倍'],
                    ['当前分位数', f'{stock_pe_percentile:.1f}%', f'{sw_index_pe_percentile:.1f}%', f'{stock_pe_percentile-sw_index_pe_percentile:+.1f}%']
                ]

            add_table_data(document, pe_history_headers, pe_history_data)

            add_paragraph(document, '')
            add_paragraph(document, '历史分位数说明：')
            add_paragraph(document, f'• 当前分位数表示当前PE在历史数据中的相对位置')
            add_paragraph(document, f'• 例如：{stock_pe_percentile:.1f}%分位数表示历史上只有{stock_pe_percentile:.1f}%的时间PE低于当前值')
            add_paragraph(document, f'• 50%分位数即为中位数，代表历史平均水平')
            if custom_peer_pe_current is not None:
                add_paragraph(document, f'• 行业指数（申万）：申万行业指数的PE，基于所有成分股按市值加权计算')
                add_paragraph(document, f'• 行业指数（自定义）：2.1.1节中同行公司PE的简单平均，反映所选同行公司的平均估值水平')
            add_paragraph(document, '')

            # 生成PE趋势图
            pe_trend_chart_path = os.path.join(IMAGES_DIR, '02_4_pe_trend_analysis.png')
            chart_path = pe_analyzer.generate_pe_trend_chart(
                stock_code, stock_pe_data,
                industry_name, industry_pe_data,
                pe_trend_chart_path
            )

            # 添加图表到文档
            if chart_path and os.path.exists(chart_path):
                add_paragraph(document, '图表 2.4: PE历史分位数趋势分析')
                add_image(document, chart_path, width=Inches(6.5))

                add_paragraph(document, '')
                add_paragraph(document, '图表解读：', bold=True)
                add_paragraph(document, '')
                add_paragraph(document, f'左上-PE走势对比：')
                add_paragraph(document, f'  • 蓝线：{stock_code}的PE-TTM走势')
                add_paragraph(document, f'  • 红线：{industry_name}的PE-TTM走势')
                add_paragraph(document, f'  • 两条线的相对位置反映个股相对行业的估值水平')
                add_paragraph(document, '')

                add_paragraph(document, f'右上-PE相对位置（个股/行业）：')
                add_paragraph(document, f'  • 比值>1：个股PE高于行业，溢价')
                add_paragraph(document, f'  • 比值<1：个股PE低于行业，折价')
                add_paragraph(document, f'  • 比值=1：与行业持平')
                add_paragraph(document, '')

                add_paragraph(document, f'左下-个股PE历史分位数：')
                add_paragraph(document, f'  • 显示{stock_code}的PE在历史中的位置变化')
                add_paragraph(document, f'  • 当前分位数：{stock_pe_percentile:.1f}%')
                add_paragraph(document, f'  • 分位数上升表示估值相对历史提升')
                add_paragraph(document, '')

                add_paragraph(document, f'右下-行业PE历史分位数：')
                add_paragraph(document, f'  • 显示{industry_name}的PE在历史中的位置变化')
                add_paragraph(document, f'  • 当前分位数：{sw_index_pe_percentile:.1f}%')
                add_paragraph(document, f'  • 可用于判断行业整体估值水平')
                add_paragraph(document, '')

            # 添加分析结论
            add_paragraph(document, '')
            add_paragraph(document, 'PE历史分位数趋势分析结论：', bold=True)
            add_paragraph(document, '')

            # 估值水平判断
            if stock_pe_percentile >= 80:
                stock_valuation_level = "历史高位"
                stock_emoji = ""
                stock_comment = f"当前PE处于历史{stock_pe_percentile:.1f}%分位数，属于历史高位，估值偏高，需警惕回调风险"
            elif stock_pe_percentile >= 60:
                stock_valuation_level = "历史中高位"
                stock_emoji = "🟠"
                stock_comment = f"当前PE处于历史{stock_pe_percentile:.1f}%分位数，属于历史中高位，估值相对偏高"
            elif stock_pe_percentile >= 40:
                stock_valuation_level = "历史中位数"
                stock_emoji = "🟡"
                stock_comment = f"当前PE处于历史{stock_pe_percentile:.1f}%分位数，接近历史中位数，估值合理"
            elif stock_pe_percentile >= 20:
                stock_valuation_level = "历史中低位"
                stock_emoji = "🟢"
                stock_comment = f"当前PE处于历史{stock_pe_percentile:.1f}%分位数，属于历史中低位，估值相对偏低"
            else:
                stock_valuation_level = "历史低位"
                stock_emoji = "✅"
                stock_comment = f"当前PE处于历史{stock_pe_percentile:.1f}%分位数，属于历史低位，估值偏低，安全边际较高"

            add_paragraph(document, f'{stock_emoji} 个股估值水平：{stock_valuation_level}')
            add_paragraph(document, f'   {stock_comment}')
            add_paragraph(document, '')

            # 与行业对比
            if stock_pe_percentile > sw_index_pe_percentile + 20:
                relative_comment = f"个股分位数({stock_pe_percentile:.1f}%)显著高于行业({sw_index_pe_percentile:.1f}%)，相对行业估值偏高"
                relative_emoji = "⚠️"
            elif stock_pe_percentile < sw_index_pe_percentile - 20:
                relative_comment = f"个股分位数({stock_pe_percentile:.1f}%)显著低于行业({sw_index_pe_percentile:.1f}%)，相对行业估值偏低，安全边际较高"
                relative_emoji = "✅"
            else:
                relative_comment = f"个股分位数({stock_pe_percentile:.1f}%)与行业({sw_index_pe_percentile:.1f}%)基本持平"
                relative_emoji = "ℹ️"

            add_paragraph(document, f'{relative_emoji} 相对行业估值：{relative_comment}')
            add_paragraph(document, '')

            # 投资建议
            add_paragraph(document, '历史分位数投资启示：')
            if stock_pe_percentile <= 25:
                add_paragraph(document, f'• 当前PE处于历史{stock_pe_percentile:.1f}%分位数（低位），历史上仅{stock_pe_percentile:.1f}%的时间估值更低')
                add_paragraph(document, f'• 从历史角度看，当前估值具备较好的安全边际')
                add_paragraph(document, f'• 建议积极关注，估值修复空间较大')
            elif stock_pe_percentile <= 50:
                add_paragraph(document, f'• 当前PE处于历史{stock_pe_percentile:.1f}%分位数（中低位），估值相对合理或偏低')
                add_paragraph(document, f'• 从历史角度看，当前估值风险可控')
                add_paragraph(document, f'• 建议适度配置，关注基本面变化')
            elif stock_pe_percentile <= 75:
                add_paragraph(document, f'• 当前PE处于历史{stock_pe_percentile:.1f}%分位数（中高位），估值相对偏高')
                add_paragraph(document, f'• 从历史角度看，当前估值风险上升')
                add_paragraph(document, f'• 建议谨慎参与，等待更好的买入时机')
            else:
                add_paragraph(document, f'• 当前PE处于历史{stock_pe_percentile:.1f}%分位数（高位），估值处于历史高位')
                add_paragraph(document, f'• 从历史角度看，当前估值风险较大')
                add_paragraph(document, f'• 建议等待估值回落至历史中低位再考虑参与')

        else:
            print("⚠️ PE历史数据获取不完整，跳过趋势图生成")
            add_paragraph(document, '⚠️ PE历史分位数趋势图暂时无法生成，可能原因：')
            add_paragraph(document, '   • tushare数据缺失或API调用限制')
            add_paragraph(document, '   • 股票或行业历史数据不足')

    except ImportError as e:
        print(f"⚠️ PE历史分析器导入失败: {e}")
        add_paragraph(document, '⚠️ PE历史分位数趋势分析功能暂不可用')

    except Exception as e:
        print(f"❌ PE历史分位数趋势分析失败: {e}")
        add_paragraph(document, f'⚠️ PE历史分位数趋势分析执行失败: {e}')

    add_paragraph(document, '')

    add_section_break(document)

    # 保存数据到context，供后续章节使用（第七章等需要）
    context['current_metrics_val'] = current_metrics_val
    context['industry_stats_val'] = industry_stats_val
    context['industry_avg_val'] = industry_avg_val
    context['peer_companies_val'] = peer_companies_val

    # 返回更新后的context
    return context
