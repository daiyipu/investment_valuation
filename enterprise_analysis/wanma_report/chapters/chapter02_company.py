#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第二章：公司基本情况
包含：公司基本信息、股本结构、公司详细信息、前十大股东、
      前十大流通股东、管理层分析、股东户数变化
"""

from typing import Dict, Any, List, Optional
import pandas as pd


class Chapter02Company:
    """公司基本情况章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any]):
        self.config = config
        self.data = data
        self.project_info = config.get('project', {})

    def generate(self) -> List[Dict[str, Any]]:
        elements = []

        elements.append({
            'type': 'heading',
            'content': '第二章 公司基本情况',
            'level': 1
        })

        # 2.1 公司基本信息
        elements.append({
            'type': 'heading',
            'content': '2.1 公司基本信息',
            'level': 2
        })
        elements.extend(self._generate_basic_info())

        # 2.2 公司详细信息
        elements.append({
            'type': 'heading',
            'content': '2.2 公司详细信息',
            'level': 2
        })
        elements.extend(self._generate_company_detail())

        # 2.3 股本结构
        elements.append({
            'type': 'heading',
            'content': '2.3 股本结构',
            'level': 2
        })
        elements.extend(self._generate_share_structure())

        # 2.4 前十大股东
        elements.append({
            'type': 'heading',
            'content': '2.4 前十大股东',
            'level': 2
        })
        elements.extend(self._generate_top10_holders())

        # 2.5 前十大流通股东
        elements.append({
            'type': 'heading',
            'content': '2.5 前十大流通股东',
            'level': 2
        })
        elements.extend(self._generate_top10_floatholders())

        # 2.6 管理层分析
        elements.append({
            'type': 'heading',
            'content': '2.6 管理层分析',
            'level': 2
        })
        elements.extend(self._generate_management())

        # 2.7 股东户数变化
        elements.append({
            'type': 'heading',
            'content': '2.7 股东户数变化',
            'level': 2
        })
        elements.extend(self._generate_holder_number())

        return elements

    def _generate_basic_info(self) -> List[Dict[str, Any]]:
        elements = []
        company_info = self.data.get('company_info', {})
        stock_code = self.project_info.get('stock_code', '')

        market_map = {
            '主板': '上海证券交易所/深圳证券交易所主板',
            '创业板': '深圳证券交易所创业板',
            '科创板': '上海证券交易所科创板',
            '北交所': '北京证券交易所',
        }
        market_name = company_info.get('market', '')
        market_desc = market_map.get(market_name, market_name)

        basic_rows = [
            {'项目': '公司名称', '内容': company_info.get('name', '')},
            {'项目': '股票代码', '内容': stock_code},
            {'项目': '上市板块', '内容': market_desc},
            {'项目': '所属行业', '内容': company_info.get('industry', '')},
            {'项目': '上市日期', '内容': self._format_date(company_info.get('list_date', ''))},
        ]

        elements.append({
            'type': 'table',
            'headers': ['项目', '内容'],
            'data': basic_rows
        })

        return elements

    def _generate_company_detail(self) -> List[Dict[str, Any]]:
        """使用stock_company接口生成公司详细信息"""
        elements = []
        stock_company = self.data.get('stock_company', {})

        if not stock_company:
            elements.append({'type': 'paragraph', 'content': '暂无公司详细信息数据。'})
            return elements

        rows = []

        # 董事长/总经理
        chairman = stock_company.get('chairman', '')
        if chairman and pd.notna(chairman):
            rows.append({'项目': '董事长', '内容': str(chairman)})
        manager = stock_company.get('manager', '')
        if manager and pd.notna(manager):
            rows.append({'项目': '总经理', '内容': str(manager)})
        legal_rep = stock_company.get('legal_rep', '')
        if legal_rep and pd.notna(legal_rep):
            rows.append({'项目': '法人代表', '内容': str(legal_rep)})

        # 注册/办公地址
        province = stock_company.get('province', '')
        city = stock_company.get('city', '')
        if province or city:
            location = f"{province}{city}"
            rows.append({'项目': '所在地区', '内容': location})

        office = stock_company.get('office', '')
        if office and pd.notna(office):
            rows.append({'项目': '办公地址', '内容': str(office)})

        # 业务范围
        business_scope = stock_company.get('business_scope', '')
        if business_scope and pd.notna(business_scope):
            scope_text = str(business_scope)
            if len(scope_text) > 200:
                scope_text = scope_text[:200] + '...'
            rows.append({'项目': '经营范围', '内容': scope_text})

        # 主营业务
        main_business = stock_company.get('main_business', '')
        if main_business and pd.notna(main_business):
            rows.append({'项目': '主营业务', '内容': str(main_business)})

        # 网站
        website = stock_company.get('website', '')
        if website and pd.notna(website):
            rows.append({'项目': '公司网站', '内容': str(website)})

        # 员工人数
        employees = stock_company.get('employees', '')
        if employees and pd.notna(employees):
            rows.append({'项目': '员工人数', '内容': f"{int(float(employees)):,} 人"})

        # 成立时间
        setup_date = stock_company.get('setup_date', '')
        if setup_date and pd.notna(setup_date):
            rows.append({'项目': '成立日期', '内容': self._format_date(str(setup_date))})

        if rows:
            elements.append({
                'type': 'table',
                'headers': ['项目', '内容'],
                'data': rows
            })

            # 公司简介
            company_info = self.data.get('company_info', {})
            name = company_info.get('name', '该公司')
            intro_parts = []
            if chairman:
                intro_parts.append(f"董事长{str(chairman)}")
            if main_business:
                intro_parts.append(f"主营业务为{str(main_business)}")
            if employees:
                intro_parts.append(f"员工{int(float(employees)):,}人")

            if intro_parts:
                elements.append({
                    'type': 'paragraph',
                    'content': f"{name}，{'，'.join(intro_parts)}。"
                })
        else:
            elements.append({'type': 'paragraph', 'content': '暂无公司详细信息。'})

        return elements

    def _generate_share_structure(self) -> List[Dict[str, Any]]:
        elements = []
        balance_sheet = self.data.get('financial_statements', {}).get('balance_sheet', pd.DataFrame())

        if balance_sheet.empty:
            elements.append({
                'type': 'paragraph',
                'content': '暂无股本结构数据。'
            })
            return elements

        # 获取最新一期
        bs_sorted = balance_sheet.sort_values('end_date', ascending=False)
        latest = bs_sorted.iloc[0]

        total_share = latest.get('total_share')
        capital_stk = latest.get('capital_stk')
        float_share = latest.get('float_share')

        rows = []
        if total_share is not None and pd.notna(total_share):
            rows.append({'项目': '总股本', '内容': f"{float(total_share):,.0f} 股"})
        if float_share is not None and pd.notna(float_share):
            rows.append({'项目': '流通股本', '内容': f"{float(float_share):,.0f} 股"})
        if capital_stk is not None and pd.notna(capital_stk):
            rows.append({'项目': '注册资本', '内容': f"{float(capital_stk):,.0f} 元"})

        if rows:
            elements.append({
                'type': 'table',
                'headers': ['项目', '内容'],
                'data': rows
            })
        else:
            elements.append({
                'type': 'paragraph',
                'content': '股本结构数据暂未获取到。'
            })

        return elements

    def _generate_top10_holders(self) -> List[Dict[str, Any]]:
        """生成前十大股东信息"""
        elements = []
        holders = self.data.get('top10_holders', pd.DataFrame())

        if holders.empty:
            elements.append({'type': 'paragraph', 'content': '暂无前十大股东数据。'})
            return elements

        # 获取最新一期
        if 'ann_date' in holders.columns:
            latest_date = holders['ann_date'].max()
            latest_holders = holders[holders['ann_date'] == latest_date]
        elif 'end_date' in holders.columns:
            latest_date = holders['end_date'].max()
            latest_holders = holders[holders['end_date'] == latest_date]
        else:
            latest_holders = holders

        table_data = []
        for rank, (_, row) in enumerate(latest_holders.iterrows(), 1):
            holder_name = row.get('holder_name', '')
            hold_amount = row.get('hold_amount', 0)
            hold_ratio = row.get('hold_ratio', 0)

            row_data = {
                '排名': rank,
                '股东名称': str(holder_name),
                '持股数量(股)': f"{float(hold_amount):,.0f}" if pd.notna(hold_amount) else '-',
                '持股比例': f"{float(hold_ratio):.2f}%" if pd.notna(hold_ratio) else '-',
            }
            table_data.append(row_data)

        if table_data:
            elements.append({
                'type': 'table',
                'headers': ['排名', '股东名称', '持股数量(股)', '持股比例'],
                'data': table_data
            })

            # 文字分析
            if len(table_data) > 0:
                top1 = table_data[0]
                top1_name = top1['股东名称']
                top1_ratio = top1['持股比例']

                # 计算前三大股东合计持股
                top3_ratio_sum = 0
                for item in table_data[:3]:
                    ratio_str = item['持股比例']
                    if ratio_str != '-':
                        top3_ratio_sum += float(ratio_str.rstrip('%'))

                text = f"公司第一大股东为{top1_name}，持股比例{top1_ratio}。"
                if top3_ratio_sum > 0:
                    text += f"前三大股东合计持股{top3_ratio_sum:.2f}%。"
                    if top3_ratio_sum > 50:
                        text += "股权较为集中。"
                    elif top3_ratio_sum > 30:
                        text += "股权相对集中。"
                    else:
                        text += "股权较为分散。"
                elements.append({'type': 'paragraph', 'content': text})

        return elements

    def _generate_top10_floatholders(self) -> List[Dict[str, Any]]:
        """生成前十大流通股东信息"""
        elements = []
        holders = self.data.get('top10_floatholders', pd.DataFrame())

        if holders.empty:
            elements.append({'type': 'paragraph', 'content': '暂无前十大流通股东数据。'})
            return elements

        # 获取最新一期
        if 'ann_date' in holders.columns:
            latest_date = holders['ann_date'].max()
            latest_holders = holders[holders['ann_date'] == latest_date]
        elif 'end_date' in holders.columns:
            latest_date = holders['end_date'].max()
            latest_holders = holders[holders['end_date'] == latest_date]
        else:
            latest_holders = holders

        table_data = []
        for rank, (_, row) in enumerate(latest_holders.iterrows(), 1):
            holder_name = row.get('holder_name', '')
            hold_amount = row.get('hold_amount', 0)
            hold_ratio = row.get('hold_ratio', 0)

            row_data = {
                '排名': rank,
                '股东名称': str(holder_name),
                '持股数量(股)': f"{float(hold_amount):,.0f}" if pd.notna(hold_amount) else '-',
                '持股比例': f"{float(hold_ratio):.2f}%" if pd.notna(hold_ratio) else '-',
            }
            table_data.append(row_data)

        if table_data:
            elements.append({
                'type': 'table',
                'headers': ['排名', '股东名称', '持股数量(股)', '持股比例'],
                'data': table_data
            })

            # 分析机构投资者占比
            institution_keywords = ['基金', '资管', '信托', '证券', '保险', '社保', 'QFII', '香港', '公司']
            inst_count = 0
            for item in table_data:
                name = item['股东名称']
                if any(kw in name for kw in institution_keywords):
                    inst_count += 1

            if inst_count > 0:
                text = f"前十大流通股东中，机构投资者占{inst_count}席，"
                if inst_count >= 5:
                    text += "流通股东以机构为主，市场对公司认可度较高。"
                elif inst_count >= 3:
                    text += "机构和自然人股东均有参与。"
                else:
                    text += "流通股东以自然人为主。"
                elements.append({'type': 'paragraph', 'content': text})

        return elements

    def _generate_management(self) -> List[Dict[str, Any]]:
        """生成管理层分析"""
        elements = []
        managers = self.data.get('stk_managers', pd.DataFrame())
        rewards = self.data.get('stk_rewards', pd.DataFrame())

        if managers.empty:
            elements.append({'type': 'paragraph', 'content': '暂无管理层信息。'})
            return elements

        # 2.6.1 核心管理团队
        elements.append({
            'type': 'heading',
            'content': '2.6.1 核心管理团队',
            'level': 3
        })

        # 筛选主要岗位
        key_positions = ['董事长', '总经理', '副总经理', '董事会秘书', '财务总监', '监事会主席']
        key_managers = managers[managers['title'].isin(key_positions)] if 'title' in managers.columns else managers

        table_data = []
        seen_names = set()
        for _, row in key_managers.iterrows():
            name = str(row.get('name', ''))
            if name in seen_names:
                continue
            seen_names.add(name)

            title = str(row.get('title', ''))
            gender = str(row.get('gender', ''))
            edu = str(row.get('edu', ''))
            resume = str(row.get('resume', ''))

            # 简历截取前100字
            resume_short = resume[:100] + '...' if len(resume) > 100 else resume

            table_data.append({
                '姓名': name,
                '职务': title,
                '性别': gender if gender and gender != 'nan' else '-',
                '学历': edu if edu and edu != 'nan' else '-',
            })

        if table_data:
            elements.append({
                'type': 'table',
                'headers': ['姓名', '职务', '性别', '学历'],
                'data': table_data
            })

        # 2.6.2 管理层薪酬（如有数据）
        if not rewards.empty:
            elements.append({
                'type': 'heading',
                'content': '2.6.2 管理层薪酬',
                'level': 3
            })

            reward_data = []
            for _, row in rewards.head(10).iterrows():
                name = str(row.get('name', ''))
                title = str(row.get('title', ''))
                salary = row.get('salary', None)

                reward_data.append({
                    '姓名': name,
                    '职务': title,
                    '薪酬(万元)': f"{float(salary):,.2f}" if pd.notna(salary) and salary else '-',
                })

            if reward_data:
                elements.append({
                    'type': 'table',
                    'headers': ['姓名', '职务', '薪酬(万元)'],
                    'data': reward_data
                })

        return elements

    def _generate_holder_number(self) -> List[Dict[str, Any]]:
        """生成股东户数变化分析"""
        elements = []
        holder_num = self.data.get('stk_holdernumber', pd.DataFrame())

        if holder_num.empty:
            elements.append({'type': 'paragraph', 'content': '暂无股东户数数据。'})
            return elements

        # 按报告期排序，取最近5期
        if 'ann_date' in holder_num.columns:
            holder_num = holder_num.sort_values('ann_date', ascending=False)
        elif 'end_date' in holder_num.columns:
            holder_num = holder_num.sort_values('end_date', ascending=False)

        recent = holder_num.head(5)

        table_data = []
        for _, row in recent.iterrows():
            end_date = row.get('end_date', '')
            ann_date = row.get('ann_date', '')
            holder_num_val = row.get('holder_num', None)

            label = str(end_date) if end_date else str(ann_date)
            if len(label) >= 8:
                label = f"{label[:4]}-{label[4:6]}-{label[6:8]}"

            table_data.append({
                '报告期': label,
                '股东户数': f"{int(float(holder_num_val)):,}" if pd.notna(holder_num_val) else '-',
            })

        if table_data:
            elements.append({
                'type': 'table',
                'headers': ['报告期', '股东户数'],
                'data': table_data
            })

            # 趋势分析
            if len(table_data) >= 2:
                try:
                    latest_num = int(table_data[0]['股东户数'].replace(',', ''))
                    prev_num = int(table_data[1]['股东户数'].replace(',', ''))
                    change = (latest_num - prev_num) / prev_num * 100
                    direction = '增加' if change > 0 else '减少'
                    text = f"最新股东户数{latest_num:,}户，环比{direction}{abs(change):.2f}%。"
                    if change > 10:
                        text += "股东户数大幅增加，筹码趋于分散。"
                    elif change < -10:
                        text += "股东户数大幅减少，筹码趋于集中。"
                    elements.append({'type': 'paragraph', 'content': text})
                except (ValueError, ZeroDivisionError):
                    pass

        return elements

    def _format_date(self, date_str: str) -> str:
        if not date_str or len(str(date_str)) < 8:
            return str(date_str)
        s = str(date_str)
        return f"{s[:4]}-{s[4:6]}-{s[6:8]}"
