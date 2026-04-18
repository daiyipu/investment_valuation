#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第三章：公司业务模式
使用Tushare fina_mainbz接口获取业务分类数据
包含：主营业务概述、业务构成分析（按产品）、业务区域分析
"""

from typing import Dict, Any, List
import pandas as pd


class Chapter03Business:
    """公司业务模式章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any]):
        self.config = config
        self.data = data
        self.project_info = config.get('project', {})

    def generate(self) -> List[Dict[str, Any]]:
        elements = []

        elements.append({
            'type': 'heading',
            'content': '第三章 公司业务模式',
            'level': 1
        })

        # 3.1 主营业务概述
        elements.append({
            'type': 'heading',
            'content': '3.1 主营业务概述',
            'level': 2
        })
        elements.extend(self._generate_business_overview())

        # 3.2 业务构成分析（按产品）
        elements.append({
            'type': 'heading',
            'content': '3.2 业务构成分析',
            'level': 2
        })
        elements.extend(self._generate_product_breakdown())

        # 3.3 业务区域分析
        elements.append({
            'type': 'heading',
            'content': '3.3 业务区域分析',
            'level': 2
        })
        elements.extend(self._generate_region_breakdown())

        return elements

    def _generate_business_overview(self) -> List[Dict[str, Any]]:
        elements = []
        company_info = self.data.get('company_info', {})
        main_business = self.data.get('main_business', pd.DataFrame())
        stock_company = self.data.get('stock_company', {})

        name = company_info.get('name', '该公司')
        industry = company_info.get('industry', '相关')

        # 主营业务描述（优先使用stock_company的main_business）
        biz_desc = ''
        if stock_company and pd.notna(stock_company.get('main_business', '')):
            biz_desc = str(stock_company['main_business'])

        if not main_business.empty and 'bz_item' in main_business.columns:
            latest_items = self._get_latest_data(main_business)
            biz_names = '、'.join(latest_items['bz_item'].dropna().unique().tolist()[:5])
            overview = f"{name}主营业务涵盖{biz_names}等领域，属于{industry}行业。"
        else:
            overview = f"{name}主营业务属于{industry}行业。"

        elements.append({'type': 'paragraph', 'content': overview})

        # 经营范围补充说明
        if biz_desc:
            elements.append({'type': 'paragraph', 'content': f"主营业务详情：{biz_desc}"})

        return elements

    def _generate_product_breakdown(self) -> List[Dict[str, Any]]:
        elements = []
        main_business = self.data.get('main_business', pd.DataFrame())

        if main_business.empty or 'bz_item' not in main_business.columns:
            elements.append({
                'type': 'paragraph',
                'content': '暂无业务构成数据。'
            })
            return elements

        # 获取最近两期数据做对比
        if 'end_date' not in main_business.columns:
            elements.append({
                'type': 'paragraph',
                'content': '业务构成数据格式异常。'
            })
            return elements

        dates = sorted(main_business['end_date'].unique(), reverse=True)

        # 最新一期
        latest_data = main_business[main_business['end_date'] == dates[0]]
        total_sales = latest_data['bz_sales'].sum()

        table_data = []
        for _, row in latest_data.iterrows():
            item_name = row.get('bz_item', '')
            sales = row.get('bz_sales', 0)
            ratio = (sales / total_sales * 100) if total_sales and total_sales > 0 else 0
            cost = row.get('bz_cost', None)
            gross_profit = sales - cost if cost and pd.notna(cost) else None
            gross_margin = (gross_profit / sales * 100) if gross_profit and sales else None

            row_data = {
                '业务板块': item_name,
                '营业收入(万元)': f"{sales / 10000:,.2f}" if pd.notna(sales) else '-',
                '占比': f"{ratio:.2f}%",
            }
            if gross_margin is not None:
                row_data['毛利率'] = f"{gross_margin:.2f}%"
            table_data.append(row_data)

        elements.append({
            'type': 'table',
            'headers': list(table_data[0].keys()) if table_data else [],
            'data': table_data
        })

        # 文字分析
        if table_data:
            top_items = sorted(table_data, key=lambda x: float(x['占比'].rstrip('%')), reverse=True)
            top3 = '、'.join([item['业务板块'] for item in top_items[:3]])
            top1_pct = top_items[0]['占比']
            elements.append({
                'type': 'paragraph',
                'content': f"公司主营业务以{top_items[0]['业务板块']}为核心（占比{top1_pct}），"
                           f"主要业务板块包括{top3}等。"
            })

        # 多期趋势（如果有两年以上数据）
        if len(dates) >= 2:
            elements.extend(self._generate_revenue_trend(main_business, dates))

        return elements

    def _generate_revenue_trend(self, main_business: pd.DataFrame, dates: list) -> List[Dict[str, Any]]:
        """生成多期收入趋势对比"""
        elements = []

        elements.append({
            'type': 'heading',
            'content': '3.2.1 业务收入趋势',
            'level': 3
        })

        periods = dates[:3]  # 最多3年
        trend_data = []

        for period in periods:
            period_data = main_business[main_business['end_date'] == period]
            period_label = str(period)[:4] + '年' if len(str(period)) >= 4 else str(period)
            total = period_data['bz_sales'].sum() / 10000 if not period_data.empty else 0

            row = {'期间': period_label, '营业收入合计(万元)': f"{total:,.2f}"}

            # 各业务板块
            for _, r in period_data.iterrows():
                item = r.get('bz_item', '')
                sales_wan = r.get('bz_sales', 0) / 10000
                col_name = f"{item}(万元)"
                row[col_name] = f"{sales_wan:,.2f}"

            trend_data.append(row)

        if trend_data:
            headers = list(trend_data[0].keys())
            elements.append({
                'type': 'table',
                'headers': headers,
                'data': trend_data
            })

        return elements

    def _generate_region_breakdown(self) -> List[Dict[str, Any]]:
        """业务区域分析（使用fina_mainbz type='R'数据）"""
        elements = []
        region_data = self.data.get('region_business', pd.DataFrame())

        if region_data.empty or 'bz_item' not in region_data.columns:
            elements.append({
                'type': 'paragraph',
                'content': '暂无区域分布数据。'
            })
            return elements

        # 获取最新一期
        if 'end_date' in region_data.columns:
            dates = sorted(region_data['end_date'].unique(), reverse=True)
            latest = region_data[region_data['end_date'] == dates[0]]
        else:
            latest = region_data

        total_sales = latest['bz_sales'].sum()
        if not total_sales or total_sales <= 0:
            elements.append({'type': 'paragraph', 'content': '区域分布数据异常。'})
            return elements

        table_data = []
        for _, row in latest.iterrows():
            region = row.get('bz_item', '')
            sales = row.get('bz_sales', 0)
            ratio = (sales / total_sales * 100) if total_sales > 0 else 0

            table_data.append({
                '区域': region,
                '营业收入(万元)': f"{sales / 10000:,.2f}" if pd.notna(sales) else '-',
                '占比': f"{ratio:.2f}%",
            })

        # 按占比降序排列
        table_data.sort(key=lambda x: float(x['占比'].rstrip('%')), reverse=True)

        if table_data:
            elements.append({
                'type': 'table',
                'headers': ['区域', '营业收入(万元)', '占比'],
                'data': table_data
            })

            # 文字分析
            top_regions = table_data[:3]
            top_region_names = '、'.join([r['区域'] for r in top_regions])
            top1_region = top_regions[0]['区域']
            top1_ratio = top_regions[0]['占比']

            text = f"公司收入主要来源于{top1_region}（占比{top1_ratio}），"
            text += f"主要收入区域包括{top_region_names}。"

            # 判断集中度
            top3_ratio_sum = sum(float(r['占比'].rstrip('%')) for r in top_regions)
            if top3_ratio_sum > 80:
                text += "区域集中度较高，需关注区域经济波动风险。"
            elif top3_ratio_sum > 60:
                text += "区域分布相对集中。"
            else:
                text += "区域分布较为均衡。"

            elements.append({'type': 'paragraph', 'content': text})

        return elements

    def _get_latest_data(self, df: pd.DataFrame) -> pd.DataFrame:
        if 'end_date' in df.columns:
            latest_date = df['end_date'].max()
            return df[df['end_date'] == latest_date]
        return df
