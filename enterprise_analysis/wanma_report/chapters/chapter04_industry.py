#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第四章：行业研究
基于东方财富研报 + LLM生成的完整行业分析
"""

from typing import Dict, Any, List
import pandas as pd
import numpy as np


class Chapter04Industry:
    """行业研究章节生成器"""

    def __init__(self, config: Dict[str, Any], data: Dict[str, Any]):
        self.config = config
        self.data = data
        self.project_info = config.get('project', {})

    def generate(self) -> List[Dict[str, Any]]:
        elements = []

        elements.append({
            'type': 'heading',
            'content': '第四章 行业研究',
            'level': 1
        })

        # 尝试LLM生成
        llm_content = self._get_llm_content()

        if llm_content:
            elements.extend(self._generate_section_4_1(llm_content))
            elements.extend(self._generate_section_4_2(llm_content))
            elements.extend(self._generate_section_4_3(llm_content))
        else:
            elements.extend(self._generate_fallback())

        # 始终生成同行对比表
        elements.extend(self._build_peer_comparison_table())

        # 始终生成研报参考列表
        elements.extend(self._generate_report_references())

        return elements

    # ----------------------------------------------------------
    # LLM调用
    # ----------------------------------------------------------

    def _get_llm_content(self) -> Dict[str, Any]:
        """尝试通过LLM生成行业分析内容"""
        llm_config = self.config.get('llm', {})

        if not llm_config:
            return None

        # 先构造LLMWriter，让它从config或环境变量获取api_key
        from utils.llm_writer import LLMWriter
        writer = LLMWriter(llm_config)

        if not writer.api_key:
            return None

        research_data = self.data.get('research_reports', {})

        # report_fetcher已经合并好文本，直接使用
        report_text = research_data.get('report_text', '')

        # 构建同行数据摘要
        peer_summary = self._build_peer_summary()

        company_info = self.data.get('company_info', {})
        company_name = company_info.get('name', self.project_info.get('name', ''))
        stock_code = self.project_info.get('stock_code', '')

        # 使用SW三级行业信息（优先）或stock_basic的行业名
        sw_info = self.data.get('sw_industry', {})
        sw_l1_name = sw_info.get('l1_name', '')
        sw_l2_name = sw_info.get('l2_name', '')
        sw_l3_name = sw_info.get('l3_name', '')
        # 构建完整行业标签：L1 > L2 > L3
        if sw_l3_name:
            industry_label = f'{sw_l1_name} > {sw_l2_name} > {sw_l3_name}'
        elif sw_l1_name:
            industry_label = sw_l1_name
        else:
            industry_label = company_info.get('industry', '')

        return writer.generate_industry_analysis(
            company_name=company_name,
            stock_code=stock_code,
            industry_name=sw_l3_name or company_info.get('industry', ''),
            sw_l1_name=industry_label,
            report_text=report_text,
            peer_data_summary=peer_summary,
        )

    def _build_peer_summary(self) -> str:
        """构建同行数据文本摘要供LLM参考"""
        industry_data = self.data.get('industry_data', pd.DataFrame())
        if industry_data.empty:
            return ''

        # 取最新年报数据
        df = industry_data.copy()
        if 'end_date' in df.columns:
            df = df[df['end_date'].astype(str).str.contains('1231', na=False)]
            df = df.sort_values('end_date', ascending=False)
            df = df.groupby('ts_code').first().reset_index()

        lines = [f"同行业主要上市公司共{df['ts_code'].nunique()}家，最新年报关键指标："]
        for _, row in df.head(10).iterrows():
            code = row.get('ts_code', '')
            roe = row.get('roe', None)
            gm = row.get('grossprofit_margin', None)
            nm = row.get('netprofit_margin', None)
            rev_yoy = row.get('or_yoy', None)

            parts = [f"- {code}"]
            if pd.notna(roe):
                parts.append(f"ROE {roe:.1f}%")
            if pd.notna(gm):
                parts.append(f"毛利率 {gm:.1f}%")
            if pd.notna(nm):
                parts.append(f"净利率 {nm:.1f}%")
            if pd.notna(rev_yoy):
                parts.append(f"营收同比 {rev_yoy:.1f}%")
            lines.append(" ".join(parts))

        return '\n'.join(lines)

    # ----------------------------------------------------------
    # LLM生成内容的三节
    # ----------------------------------------------------------

    def _generate_section_4_1(self, llm_content: dict) -> List[Dict]:
        elements = [{'type': 'heading', 'content': '4.1 行业市场概述', 'level': 2}]

        s1 = llm_content.get('section_4_1', {})

        for key, title in [
            ('intro', '4.1.1 行业介绍'),
            ('supply_chain', '4.1.2 产业链分析'),
            ('regulation', '4.1.3 行业监管与政策'),
            ('technology', '4.1.4 技术发展趋势'),
        ]:
            text = s1.get(key, '')
            if text:
                elements.append({'type': 'heading', 'content': title, 'level': 3})
                elements.append({'type': 'paragraph', 'content': text})

        return elements

    def _generate_section_4_2(self, llm_content: dict) -> List[Dict]:
        elements = [{'type': 'heading', 'content': '4.2 行业市场规模', 'level': 2}]

        s2 = llm_content.get('section_4_2', {})

        for key, title in [
            ('drivers', '4.2.1 市场驱动因素'),
            ('global_market', '4.2.2 全球市场规模'),
            ('domestic_market', '4.2.3 国内市场规模及趋势'),
        ]:
            text = s2.get(key, '')
            if text:
                elements.append({'type': 'heading', 'content': title, 'level': 3})
                elements.append({'type': 'paragraph', 'content': text})

        return elements

    def _generate_section_4_3(self, llm_content: dict) -> List[Dict]:
        elements = [{'type': 'heading', 'content': '4.3 行业竞争格局', 'level': 2}]

        s3 = llm_content.get('section_4_3', {})

        for key, title in [
            ('landscape', '4.3.1 竞争格局概述'),
            ('competitors', '4.3.2 主要竞争对手'),
        ]:
            text = s3.get(key, '')
            if text:
                elements.append({'type': 'heading', 'content': title, 'level': 3})
                elements.append({'type': 'paragraph', 'content': text})

        return elements

    # ----------------------------------------------------------
    # 同行对比表（4.3.3，纯Tushare数据）
    # ----------------------------------------------------------

    def _build_peer_comparison_table(self) -> List[Dict]:
        elements = [{'type': 'heading', 'content': '4.3.3 上市公司主要情况对比', 'level': 3}]

        industry_data = self.data.get('industry_data', pd.DataFrame())
        if industry_data.empty:
            elements.append({'type': 'paragraph', 'content': '暂无同行对比数据。'})
            return elements

        df = industry_data.copy()
        if 'end_date' in df.columns:
            df = df[df['end_date'].astype(str).str.contains('1231', na=False)]
            df = df.sort_values('end_date', ascending=False)
            df = df.groupby('ts_code').first().reset_index()

        # 获取公司名称映射
        name_map = self._get_peer_names(df['ts_code'].unique().tolist())
        target_code = self.project_info.get('stock_code', '')

        table_data = []
        for _, row in df.iterrows():
            code = row.get('ts_code', '')
            name = name_map.get(code, code)

            table_data.append({
                '公司名称': f"{'★ ' if code == target_code else ''}{name}",
                'ROE(%)': self._fmt(row.get('roe')),
                '毛利率(%)': self._fmt(row.get('grossprofit_margin')),
                '净利率(%)': self._fmt(row.get('netprofit_margin')),
                '营收同比(%)': self._fmt(row.get('or_yoy')),
                '净利润同比(%)': self._fmt(row.get('netprofit_yoy')),
                '资产负债率(%)': self._fmt(row.get('debt_to_assets')),
            })

        # 目标公司排第一
        table_data.sort(key=lambda x: (0 if '★' in x['公司名称'] else 1, x['公司名称']))

        if table_data:
            elements.append({
                'type': 'table',
                'headers': list(table_data[0].keys()),
                'data': table_data,
            })

        return elements

    # ----------------------------------------------------------
    # 研报参考列表（4.4）
    # ----------------------------------------------------------

    def _generate_report_references(self) -> List[Dict]:
        elements = [{'type': 'heading', 'content': '4.4 行业研究报告参考', 'level': 2}]

        research_data = self.data.get('research_reports', {})
        reports = research_data.get('reports', [])

        if not reports:
            elements.append({'type': 'paragraph', 'content': '暂无行业研究报告数据。'})
            return elements

        table_data = []
        for i, rep in enumerate(reports, 1):
            pub_date = str(rep.get('publishDate', ''))[:10]
            web_url = rep.get('web_url', '')
            table_data.append({
                '序号': str(i),
                '报告标题': rep.get('title', ''),
                '研究机构': rep.get('orgName', ''),
                '发布日期': pub_date,
                '行业': rep.get('industryName', ''),
                '链接': web_url if web_url else '-',
            })

        elements.append({
            'type': 'table',
            'headers': ['序号', '报告标题', '研究机构', '发布日期', '行业', '链接'],
            'data': table_data,
        })

        return elements

    # ----------------------------------------------------------
    # 回退模式（无LLM/无研报时使用）
    # ----------------------------------------------------------

    def _generate_fallback(self) -> List[Dict]:
        """基于Tushare数据的基础行业分析"""
        elements = []
        company_info = self.data.get('company_info', {})
        industry_name = company_info.get('industry', '相关')

        # 4.1 行业概述（基础版）
        elements.append({'type': 'heading', 'content': '4.1 行业市场概述', 'level': 2})
        elements.append({
            'type': 'paragraph',
            'content': (
                f"公司所属{industry_name}行业。以下行业分析基于Tushare公开财务数据进行量化分析，"
                "详细的行业研究报告内容需配置LLM接口和研报获取功能后自动生成。"
            ),
        })

        # 行业关键指标统计
        industry_data = self.data.get('industry_data', pd.DataFrame())
        if not industry_data.empty:
            df = industry_data.copy()
            if 'end_date' in df.columns:
                df = df[df['end_date'].astype(str).str.contains('1231', na=False)]
                df = df.sort_values('end_date', ascending=False)
                df = df.groupby('ts_code').first().reset_index()

            stats = []
            for field, label in [
                ('roe', 'ROE'), ('grossprofit_margin', '毛利率'),
                ('netprofit_margin', '净利率'), ('current_ratio', '流动比率'),
                ('debt_to_assets', '资产负债率'),
            ]:
                if field in df.columns:
                    vals = df[field].dropna()
                    if len(vals) > 0:
                        stats.append({
                            '指标': label,
                            '行业中位数': f"{vals.median():.2f}",
                            '行业平均': f"{vals.mean():.2f}",
                            '行业最大': f"{vals.max():.2f}",
                            '行业最小': f"{vals.min():.2f}",
                            '公司数量': str(len(vals)),
                        })

            if stats:
                elements.append({
                    'type': 'heading',
                    'content': '4.1.1 行业关键指标统计',
                    'level': 3,
                })
                elements.append({
                    'type': 'table',
                    'headers': list(stats[0].keys()),
                    'data': stats,
                })

        # 4.2 市场规模（占位）
        elements.append({'type': 'heading', 'content': '4.2 行业市场规模', 'level': 2})
        elements.append({
            'type': 'paragraph',
            'content': f"{industry_name}行业的市场规模分析需结合行业研报数据，待配置LLM接口后自动生成详细内容。",
        })

        # 4.3 竞争格局标题（表格在generate()中始终生成）
        elements.append({'type': 'heading', 'content': '4.3 行业竞争格局', 'level': 2})

        return elements

    # ----------------------------------------------------------
    # 辅助方法
    # ----------------------------------------------------------

    def _get_peer_names(self, ts_codes: list) -> dict:
        """获取同行公司名称映射"""
        name_map = {}
        # 先从已有的industry_data中尝试获取
        industry_data = self.data.get('industry_data', pd.DataFrame())
        if not industry_data.empty and 'ts_code' in industry_data.columns:
            # fina_indicator中没有公司名称，需要从stock_basic获取
            pass

        # 通过Tushare获取
        try:
            import tushare as ts
            token = self.config.get('tushare', {}).get('token', '')
            if not token:
                import os
                token = os.environ.get('TUSHARE_TOKEN', '')
            if token:
                pro = ts.pro_api(token)
                target_code = self.project_info.get('stock_code', '')
                all_codes = list(set(ts_codes + [target_code]))
                # 分批查询（每次不超过100个）
                for i in range(0, len(all_codes), 100):
                    batch = all_codes[i:i + 100]
                    df = pro.stock_basic(
                        ts_code=','.join(batch),
                        fields='ts_code,name',
                    )
                    if df is not None and not df.empty:
                        for _, row in df.iterrows():
                            name_map[row['ts_code']] = row['name']
        except Exception:
            pass

        return name_map

    def _fmt(self, value, decimal: int = 2) -> str:
        """格式化数值"""
        if value is None:
            return '-'
        try:
            v = float(value)
            if pd.isna(v):
                return '-'
            return f"{v:.{decimal}f}"
        except (ValueError, TypeError):
            return '-'
