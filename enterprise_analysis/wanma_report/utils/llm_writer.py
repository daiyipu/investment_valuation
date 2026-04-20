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
# 近期行业研报参考（正文摘要）
以下是东方财富获取的{industry_label}行业近期研究报告的正文摘要，每篇研报标注了来源机构。你必须基于这些研报内容进行摘要和引用：

{report_text}
"""
        else:
            research_section = f'\n# 近期行业研报参考\n（未获取到{industry_label}行业的研报数据，请标注"据行业公开资料"。）\n'

        peer_section = ""
        if peer_data_summary:
            peer_section = f"""
# 同行业上市公司关键指标（来自Tushare财务数据）
{peer_data_summary}
"""

        return f"""请为{company_name}（{stock_code}）撰写一份行业研究分析，用于企业基本面分析报告的第四章。

重要：该公司属于申万行业分类 "{industry_label}"，请聚焦于该细分行业进行分析，不要泛泛而谈大一级行业。

{research_section}
{peer_section}

# 核心原则（必须严格遵守）
1. **所有内容必须源自上述研报**，严禁自行编造数据或观点
2. 每一个事实、数据、观点都必须标注出处，格式为脚注标记[n]
3. 在JSON中新增"footnotes"字段，列出所有脚注来源

# 输出要求
请严格按照以下JSON格式输出，每个字段为中文段落文本（200-500字），使用客观、专业的分析语言：

{{
  "section_4_1": {{
    "intro": "行业介绍：聚焦{industry_name}细分行业的定义、主要组成部分、发展阶段、在国民经济中的地位",
    "supply_chain": "产业链分析：上游原材料/零部件供应、中游制造/集成环节、下游应用领域及需求特点",
    "regulation": "行业监管与政策：主管单位、核心监管政策、近年重要产业政策及支持方向",
    "technology": "技术发展趋势：核心技术方向、行业技术迭代趋势、主要技术壁垒和进入门槛"
  }},
  "section_4_2": {{
    "drivers": "市场驱动因素：推动行业增长的核心结构性因素分析（3-4个）",
    "global_market": "全球市场规模：全球市场规模数据、近5年复合增长率(CAGR)、增长趋势、区域分布",
    "domestic_market": "国内市场规模：中国市场规模数据及增速、CAGR、在全球市场中的占比、未来3-5年市场预测"
  }},
  "section_4_3": {{
    "landscape": "竞争格局概述：行业集中度、竞争层次（国际vs国内）、市场份额分布格局、行业进入壁垒",
    "competitors": "主要竞争对手：列举3-5家行业代表性公司，描述其核心业务、市场定位和竞争优势"
  }},
  "footnotes": {{
    "1": "出处说明（如：头豹研究院《2026年电网设备出海行业词条报告》，2026-04-20）",
    "2": "出处说明..."
  }}
}}

# 脚注规则
- 引用研报数据/观点时，在句末加[n]标记（如"市场规模达1915亿美元[1]"）
- footnotes中每条格式："机构名《报告标题》，发布日期"
- 研报中的数据直接引用，不要修改数值
- 如果某段落没有引用任何研报，说明该段落内容缺乏来源，应缩减为简短概述

# 其他注意事项
1. 涉及市场规模数据时，必须给出复合增长率(CAGR)
2. 不要包含markdown格式（如**、#等），仅输出纯文本段落
3. 必须输出合法JSON，不要在JSON前后添加其他文字
4. 分析应紧扣{industry_name}细分行业特点"""

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
