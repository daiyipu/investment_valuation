#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
东方财富研报获取
基于report_spyder项目的api_spider.py简化实现
支持PDF下载+文本提取，用于LLM行业研究内容生成
"""

import io
import json
import re
import time
from typing import List, Dict, Optional
from datetime import datetime, timedelta

import requests


class ReportFetcher:
    """东方财富研报元数据获取 + PDF文本提取"""

    def __init__(self, config: dict):
        self.config = config
        self.api_url = "https://reportapi.eastmoney.com/report/list"
        self.max_reports = config.get('max_reports', 8)
        self.max_text_per_report = config.get('max_text_per_report', 8000)
        self.max_total_text = config.get('max_total_text', 40000)
        self.pdf_timeout = config.get('pdf_timeout', 30)
        self.download_delay = config.get('download_delay', 1.0)
        self.date_range_months = config.get('date_range_months', 24)

        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Referer': 'https://data.eastmoney.com/',
            'Origin': 'https://data.eastmoney.com',
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        })

    # ----------------------------------------------------------
    # 研报元数据获取
    # ----------------------------------------------------------

    def fetch_stock_reports(self, stock_code: str, max_count: int = 5) -> List[dict]:
        """获取个股研报（API的stockCode过滤已失效，客户端过滤）"""
        code = stock_code.split('.')[0]
        begin_time = (datetime.now() - timedelta(days=self.date_range_months * 30)).strftime('%Y-%m-%d')

        params = {
            'pageSize': '100',
            'pageNo': '1',
            'qType': '0',
            'beginTime': begin_time,
            'endTime': '',
            'fields': '',
            '_': str(int(time.time() * 1000)),
        }

        try:
            r = self.session.get(self.api_url, params=params, timeout=self.pdf_timeout)
            r.raise_for_status()
            data = r.json()
            reports = data.get('data') or []
            # 客户端按stockCode过滤
            matched = [rep for rep in reports if rep.get('stockCode') == code]
            return [self._parse_meta(rep) for rep in matched[:max_count]]
        except Exception as e:
            print(f"  获取个股研报失败: {e}")
            return []

    def fetch_industry_reports(self, industry_keyword: str, max_count: int = 5) -> List[dict]:
        """获取行业研报（逐级搜索：L3精确 → L2补充 → L1兜底）"""
        begin_time = (datetime.now() - timedelta(days=self.date_range_months * 30)).strftime('%Y-%m-%d')

        tiers = self._get_tiered_keywords(industry_keyword)
        all_reports = []
        seen_codes = set()
        tier_names = ['L3', 'L2', 'L1']

        for idx, tier_keywords in enumerate(tiers):
            if not tier_keywords or len(all_reports) >= max_count:
                continue
            remaining = max_count - len(all_reports)

            reports, seen_codes = self._search_reports(tier_keywords, begin_time, remaining, seen_codes)
            tier_label = tier_names[idx] if idx < len(tier_names) else f'L{idx+1}'
            print(f'  研报搜索 {tier_label} {tier_keywords}: 匹配 {len(reports)} 篇')
            all_reports.extend(reports)

        return all_reports[:max_count]

    def _search_reports(self, keywords, begin_time, max_count, seen_codes=None) -> tuple:
        """按关键词搜索行业研报"""
        if seen_codes is None:
            seen_codes = set()
        all_reports = []
        page = 1

        while len(all_reports) < max_count and page <= 10:
            params = {
                'industryCode': '*',
                'pageSize': '50',
                'pageNo': str(page),
                'qType': '1',
                'beginTime': begin_time,
                'endTime': '',
                'fields': '',
                '_': str(int(time.time() * 1000)),
            }

            try:
                r = self.session.get(self.api_url, params=params, timeout=self.pdf_timeout)
                r.raise_for_status()
                data = r.json()
                page_data = data.get('data') or []
                if not page_data:
                    break

                for rep in page_data:
                    info_code = rep.get('infoCode', '')
                    if info_code in seen_codes:
                        continue

                    title = rep.get('title', '')
                    industry_name = rep.get('industryName', '')

                    if any(kw in title or kw in industry_name for kw in keywords):
                        seen_codes.add(info_code)
                        all_reports.append(self._parse_meta(rep))
                        if len(all_reports) >= max_count:
                            break

                page += 1
                time.sleep(0.5)
            except Exception as e:
                print(f"  获取行业研报第{page}页失败: {e}")
                break

        return all_reports, seen_codes

    @staticmethod
    def _split_industry_name(name: str) -> list:
        """从行业名称中提取核心关键词

        逐层去掉无意义后缀，每层保留有意义的结果：
        "线缆部件及其他" → ["线缆部件及其他", "线缆部件", "线缆"]
        "汽车零部件" → ["汽车零部件", "汽车"]
        "电网设备" → ["电网设备", "电网"]
        """
        keywords = [name]
        current = name
        # 按顺序逐层剥离后缀
        for suffix in ['及其他', '及其他设备', '设备', '部件', '制造', '加工', '服务', '行业', '产业']:
            if current.endswith(suffix):
                stripped = current[:-len(suffix)]
                if len(stripped) >= 2 and stripped not in keywords:
                    keywords.append(stripped)
                    current = stripped
        return keywords

    @staticmethod
    def _get_tiered_keywords(industry_name: str) -> list:
        """按SW行业层级拆分关键词，返回分层列表

        返回: [[L3关键词], [L2关键词], [L1关键词]]
        如 "传媒, 平面媒体, 大众出版"
        → [["大众出版"], ["平面媒体"], ["传媒"]]
        """
        parts = [k.strip() for k in industry_name.split(',') if k.strip()] if industry_name else []
        if not parts:
            return [[industry_name]] if industry_name else [[]]

        tiers = []
        for part in reversed(parts):  # L3, L2, L1 顺序
            tier_kw = ReportFetcher._split_industry_name(part)
            tiers.append(tier_kw)
        return tiers

    # ----------------------------------------------------------
    # 研报正文提取（从东方财富网页提取）
    # ----------------------------------------------------------

    def fetch_report_text(self, info_code: str, max_chars: int = 0) -> Optional[str]:
        """从东方财富网页提取研报正文（通过页面内嵌的zwinfo JSON）

        Args:
            info_code: 研报编号（如 AP202604201821328430）
            max_chars: 最大提取字符数，0表示使用配置默认值

        Returns:
            提取的文本内容，失败返回None
        """
        if not max_chars:
            max_chars = self.max_text_per_report

        web_url = f"https://data.eastmoney.com/report/zw_stock.jshtml?infocode={info_code}"

        try:
            resp = self.session.get(web_url, timeout=self.pdf_timeout)
            resp.raise_for_status()

            # 从页面HTML中提取zwinfo JSON
            m = re.search(r'var zwinfo=\s*(\{.*?\});', resp.text, re.DOTALL)
            if not m:
                return None

            info = json.loads(m.group(1))
            content = info.get('notice_content', '')
            if not content:
                return None

            # 清理HTML标签
            content = re.sub(r'<[^>]+>', '', content)
            content = re.sub(r'\s+', ' ', content).strip()
            return content[:max_chars]

        except Exception as e:
            print(f"  研报正文提取失败({info_code}): {e}")
            return None

    # ----------------------------------------------------------
    # 高层接口
    # ----------------------------------------------------------

    def fetch_reports_for_chapter(
        self,
        stock_code: str,
        industry_name: str,
        sw_l1_name: str = "",
    ) -> dict:
        """获取研报元数据 + 提取文本内容

        Returns:
            {
                reports: [{title, publishDate, orgName, infoCode, industryName, pdf_url, web_url}],
                report_text: str,        # 合并后的研报文本，用于LLM输入
                download_errors: list,   # 提取失败的研报infoCode列表
            }
        """
        # 1. 获取研报元数据
        stock_reports = self.fetch_stock_reports(stock_code, max_count=5)
        time.sleep(self.download_delay)

        keyword = sw_l1_name or industry_name
        industry_reports = self.fetch_industry_reports(keyword, max_count=5)

        # 合并去重
        seen = set()
        all_reports = []
        for rep in stock_reports + industry_reports:
            code = rep.get('infoCode', '')
            if code and code not in seen:
                seen.add(code)
                # 添加链接信息
                rep['pdf_url'] = f"https://pdf.dfcfw.com/pdf/H3_{code}_1.pdf"
                rep['web_url'] = f"https://data.eastmoney.com/report/zw_stock.jshtml?infocode={code}"
                all_reports.append(rep)

        all_reports.sort(key=lambda x: x.get('publishDate', ''), reverse=True)
        all_reports = all_reports[:self.max_reports]

        # 2. 从网页提取研报正文
        extracted_texts = {}
        download_errors = []

        for rep in all_reports:
            info_code = rep.get('infoCode', '')
            if not info_code:
                continue

            text = self.fetch_report_text(info_code)
            if text and len(text.strip()) > 100:
                extracted_texts[info_code] = text
                title = rep.get('title', '')
                print(f"  已提取研报文本: 《{title}》({len(text)}字)")
            else:
                download_errors.append(info_code)

            time.sleep(self.download_delay)

        # 3. 构建LLM输入文本
        report_text = self._build_report_text_for_llm(all_reports, extracted_texts)

        return {
            'reports': all_reports,
            'report_text': report_text,
            'download_errors': download_errors,
        }

    def _build_report_text_for_llm(self, reports: list, extracted_texts: dict) -> str:
        """构建LLM可用的参考文本（研报元数据 + 提取的正文摘要）"""
        if not reports:
            return ''

        lines = []
        total_chars = 0

        for rep in reports:
            info_code = rep.get('infoCode', '')
            date = str(rep.get('publishDate', ''))[:10]
            org = rep.get('orgName', '')
            title = rep.get('title', '')
            industry = rep.get('industryName', '')

            header = f"[{date}] {org}《{title}》{'(' + industry + ')' if industry else ''}"
            lines.append(header)

            # 如果有提取的正文，截取前2000字作为摘要
            text = extracted_texts.get(info_code, '')
            if text:
                excerpt = text[:2000]
                lines.append(f"摘要：{excerpt}")
            else:
                lines.append("（PDF文本提取失败，仅参考标题方向）")

            lines.append('')  # 空行分隔
            total_chars += sum(len(l) for l in lines[-3:])

            if total_chars >= self.max_total_text:
                break

        return '\n'.join(lines)

    def _build_report_text(self, reports: list) -> str:
        """将研报元数据构建为参考文本列表（用于报告4.4节）"""
        if not reports:
            return ''

        lines = []
        for rep in reports:
            date = str(rep.get('publishDate', ''))[:10]
            org = rep.get('orgName', '')
            title = rep.get('title', '')
            industry = rep.get('industryName', '')
            lines.append(f"- [{date}] {org}《{title}》{'(' + industry + ')' if industry else ''}")

        return '\n'.join(lines)

    def _parse_meta(self, api_data: dict) -> dict:
        """解析API返回的研报元数据"""
        return {
            'title': api_data.get('title', ''),
            'publishDate': api_data.get('publishDate', ''),
            'orgName': api_data.get('orgSName', '') or api_data.get('orgName', ''),
            'infoCode': api_data.get('infoCode', ''),
            'industryName': api_data.get('industryName', ''),
        }
