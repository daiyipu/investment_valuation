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
        """获取行业研报（分两轮搜索：先用精确关键词，不够再用宽泛关键词补充）"""
        begin_time = (datetime.now() - timedelta(days=self.date_range_months * 30)).strftime('%Y-%m-%d')

        # 从行业名称生成搜索关键词
        all_keywords = self._expand_keywords(industry_keyword)
        # L1级别的宽泛关键词（用于第二轮补充）
        broad_keywords = self._get_broad_keywords(industry_keyword)

        # 第一轮：精确关键词搜索
        all_reports, seen_codes = self._search_reports(all_keywords, begin_time, max_count)

        # 第二轮：如果精确搜索不够，用宽泛关键词补充
        if len(all_reports) < max_count and broad_keywords:
            remaining = max_count - len(all_reports)
            broad_reports, _ = self._search_reports(broad_keywords, begin_time, remaining, seen_codes)
            all_reports.extend(broad_reports)

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
    def _get_broad_keywords(industry_name: str) -> list:
        """提取L1级别宽泛关键词，用于第二轮补充搜索"""
        parts = [k.strip() for k in industry_name.split(',') if k.strip()]
        if not parts:
            return []

        # L1名称映射到宽泛关键词（补充搜索用）
        broad_map = {
            '电力设备': ['电力设备', '电力'],
            '电气设备': ['电力设备', '电气'],
        }
        l1_name = parts[0]
        return broad_map.get(l1_name, [])

    @staticmethod
    def _expand_keywords(industry_name: str) -> list:
        """从行业名称生成搜索关键词列表

        支持逗号分隔的SW行业层级（如"电力设备, 电网设备, 线缆部件及其他"）
        优先按最细粒度（L3）的行业名匹配，避免L1太宽泛引入不相关关键词
        """
        # 基础关键词：按逗号分割的行业名 [L1, L2, L3]
        parts = [k.strip() for k in industry_name.split(',') if k.strip()] if industry_name else []

        # 行业名称到搜索关键词的扩展映射（所有层级统一映射）
        expansion_map = {
            # === 电力设备 L1 下的 L2/L3 ===
            '线缆部件及其他': ['线缆', '电缆', '电线', '电网', '输变电', '配电'],
            '输变电设备': ['输变电', '变压器', '电网', '高压', '电缆'],
            '配电设备': ['配电', '开关', '电网', '电缆'],
            '电网自动化设备': ['电网', '自动化', '智能电网', '配电'],
            '电网设备': ['电网', '输变电', '配电', '电力', '电缆'],
            '电机': ['电机', '电动机', '发电机'],
            '光伏设备': ['光伏', '太阳能', '硅料'],
            '风电设备': ['风电', '风力'],
            '锂电池': ['锂电', '电池', '储能'],
            '光伏电池组件': ['光伏', '太阳能'],
            # === L1 级别（最宽泛，仅当L2/L3都没匹配时使用）===
            '电力设备': ['电力设备', '电力'],
            '电气设备': ['电力设备', '电气', '电力'],
            '汽车': ['汽车', '新能源车', '电动'],
            '医药': ['医药', '生物', '医疗'],
            '食品': ['食品', '饮料', '白酒'],
            '计算机': ['计算机', '软件', 'AI', '人工智能'],
            '电子': ['电子', '半导体', '芯片'],
            '基础化工': ['化工', '化学', '化纤'],
            '有色金属': ['有色', '金属', '铜', '铝'],
            '银行': ['银行', '金融', '信贷'],
            '房地产': ['房地产', '地产', '住房'],
            '通信': ['通信', '5G', '光通信'],
            '国防军工': ['军工', '国防', '航天'],
            '机械设备': ['机械', '设备', '工程机械'],
            '建筑材料': ['建材', '水泥', '玻璃'],
            '公用事业': ['公用事业', '电力', '燃气', '环保'],
            '交通运输': ['交通', '物流', '航空', '港口'],
            '传媒': ['传媒', '影视', '游戏', '广告'],
        }

        # 从最细粒度开始匹配：parts按[L1, L2, L3]排列，倒序遍历
        expanded = []
        matched_level = -1  # 匹配到的层级索引

        for i in range(len(parts) - 1, -1, -1):
            part = parts[i]
            for key, words in expansion_map.items():
                if key in part:
                    expanded = words
                    matched_level = i
                    break
            if expanded:
                break

        if not expanded:
            return list(set(parts)) if parts else [industry_name]

        # 只保留匹配层级及以下的关键词（排除更粗粒度的名称）
        # 例如L3匹配时，不包含L1（太宽泛），但保留L2和L3自身
        keywords = parts[matched_level:]  # 从匹配层级开始取
        keywords.extend(expanded)

        return list(set(keywords))

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
