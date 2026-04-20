#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LLM API封装 - 用于行业研究章节内容生成
支持OpenAI兼容接口（包括智谱GLM等）
"""

import os
import re
import json
from typing import Optional


class LLMWriter:
    """LLM API封装，用于生成行业分析内容"""

    def __init__(self, config: dict):
        self.api_key = config.get('api_key', '') or os.environ.get('LLM_API_KEY', '') or os.environ.get('ZHIPUAI_API_KEY', '')
        self.base_url = config.get('base_url', 'https://open.bigmodel.cn/api/coding/paas/v4')
        self.model = config.get('model', 'glm-4-flash')
        self.max_tokens = config.get('max_tokens', 8192)
        self.temperature = config.get('temperature', 0.3)
        self._client = None

    def _get_client(self):
        if self._client is None:
            try:
                from openai import OpenAI
            except ImportError:
                raise ImportError("openai包未安装，请运行: pip install openai")
            self._client = OpenAI(api_key=self.api_key, base_url=self.base_url)
        return self._client

    def generate_industry_analysis(
        self,
        company_name: str,
        stock_code: str,
        industry_name: str,
        sw_l1_name: str,
        report_text: str,
        peer_data_summary: str,
    ) -> Optional[dict]:
        """调用LLM生成行业分析内容

        Args:
            report_text: 研报元数据汇总（标题+机构+日期）
            peer_data_summary: Tushare同行财务数据摘要

        Returns:
            结构化分析内容dict 或 None
        """
        if not self.api_key:
            print("  LLM API key未配置，跳过LLM生成")
            return None

        prompt = self._build_prompt(
            company_name, stock_code, industry_name, sw_l1_name,
            report_text, peer_data_summary,
        )

        try:
            client = self._get_client()
            response = client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一位专业的金融行业分析师，擅长撰写行业研究报告。请严格按照要求的JSON格式输出。"},
                    {"role": "user", "content": prompt},
                ],
                max_tokens=self.max_tokens,
                temperature=self.temperature,
                timeout=180,
            )

            raw_text = response.choices[0].message.content.strip()
            result = self._parse_response(raw_text)
            if result:
                result['model_used'] = self.model
            return result

        except ImportError as e:
            print(f"  LLM依赖缺失: {e}")
            return None
        except Exception as e:
            print(f"  LLM调用失败: {e}")
            return None

    def _build_prompt(
        self, company_name, stock_code, industry_name, sw_l1_name,
        report_text, peer_data_summary,
    ) -> str:
        industry_label = sw_l1_name or industry_name

        research_section = ""
        if report_text:
            research_section = f"""
# 近期行业研报参考
以下是从东方财富获取的近期{industry_label}行业相关研究报告列表（标题+机构+日期），请结合这些研报方向和你的行业知识进行分析：

{report_text}
"""
        else:
            research_section = f"\n# 近期行业研报参考\n（未获取到{industry_label}行业的研报数据，请基于行业常识和提供的同行数据进行分析。）\n"

        peer_section = ""
        if peer_data_summary:
            peer_section = f"""
# 同行业上市公司关键指标
{peer_data_summary}
"""

        return f"""请为{company_name}（{stock_code}）所属的{industry_label}行业撰写一份行业研究分析，用于企业基本面分析报告的第四章。

{research_section}
{peer_section}

# 输出要求
请严格按照以下JSON格式输出，每个字段为中文段落文本（200-500字），使用客观、专业的分析语言：

{{
  "section_4_1": {{
    "intro": "行业介绍：包括行业定义、主要组成部分、发展阶段、在国民经济中的地位",
    "supply_chain": "产业链分析：上游原材料/零部件供应、中游制造/集成环节、下游应用领域及需求特点",
    "regulation": "行业监管与政策：主管单位、核心监管政策、近年重要产业政策及支持方向",
    "technology": "技术发展趋势：核心技术方向、行业技术迭代趋势、主要技术壁垒和进入门槛"
  }},
  "section_4_2": {{
    "drivers": "市场驱动因素：推动行业增长的核心结构性因素分析（3-4个），如政策、需求、技术等",
    "global_market": "全球市场规模：全球市场规模数据（引用具体数据和来源机构）、增长趋势、区域分布",
    "domestic_market": "国内市场规模：中国市场规模数据及增速、在全球市场中的占比、未来3-5年市场预测"
  }},
  "section_4_3": {{
    "landscape": "竞争格局概述：行业集中度、竞争层次（国际vs国内）、市场份额分布格局、行业进入壁垒",
    "competitors": "主要竞争对手：列举3-5家行业代表性公司，描述其核心业务、市场定位和竞争优势"
  }}
}}

注意事项：
1. 涉及市场规模数据时，需注明来源机构（如"据GGII"、"据前瞻产业研究院"等）
2. 如数据无法确认，请使用"据公开资料"或"据行业统计数据"
3. 不要包含markdown格式（如**、#等），仅输出纯文本段落
4. 必须输出合法JSON，不要在JSON前后添加其他文字
5. 分析应紧扣{industry_label}行业特点，内容要有深度和数据支撑"""

    def _parse_response(self, response_text: str) -> Optional[dict]:
        """解析LLM返回的JSON"""
        try:
            return json.loads(response_text)
        except json.JSONDecodeError:
            pass

        m = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', response_text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(1))
            except json.JSONDecodeError:
                pass

        start = response_text.find('{')
        end = response_text.rfind('}')
        if start >= 0 and end > start:
            try:
                return json.loads(response_text[start:end + 1])
            except json.JSONDecodeError:
                pass

        print(f"  LLM响应JSON解析失败，前200字: {response_text[:200]}")
        return None
