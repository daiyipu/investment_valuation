#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
East Money Research Report Agent (东方财富研报Agent)

Fetches industry research reports from East Money, extracts text,
and generates structured analysis using LLM.
Uses Claude Agent SDK for intelligent orchestration with fallback to direct pipeline.

Usage:
    # CLI (SDK agent mode - Claude orchestrates)
    python3 -m agents.eastmoney_research_agent --stock 002276.SZ --name 万马高分子 --industry 线缆部件及其他

    # CLI (direct pipeline mode - no SDK)
    python3 -m agents.eastmoney_research_agent --stock 002276.SZ --name 万马高分子 --industry 线缆部件及其他 --no-sdk

    # Python API
    from agents import run_research
    result = await run_research(stock_code='002276.SZ', company_name='万马高分子', industry_name='线缆部件及其他')

    # Or use the agent class directly
    from agents import EastMoneyResearchAgent
    agent = EastMoneyResearchAgent(config={...})
    result = await agent.run(stock_code='002276.SZ', company_name='万马高分子')
"""

import asyncio
import json
import os
import sys
import argparse
from typing import Optional, TypedDict

# ---------------------------------------------------------------------------
# Add project root to sys.path so utils/ can be imported from wanma_report
# ---------------------------------------------------------------------------
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_WANMA_DIR = os.path.join(_PROJECT_ROOT, 'enterprise_analysis', 'wanma_report')
if _WANMA_DIR not in sys.path:
    sys.path.insert(0, _WANMA_DIR)

from utils.report_fetcher import ReportFetcher
from utils.llm_writer import LLMWriter


# ---------------------------------------------------------------------------
# Default config
# ---------------------------------------------------------------------------
DEFAULT_FETCHER_CONFIG = {
    'max_reports': 8,
    'max_text_per_report': 8000,
    'max_total_text': 40000,
    'pdf_timeout': 30,
    'download_delay': 1.0,
    'date_range_months': 24,
}

DEFAULT_LLM_CONFIG = {
    'base_url': os.environ.get('LLM_BASE_URL', 'https://open.bigmodel.cn/api/coding/paas/v4'),
    'model': os.environ.get('LLM_MODEL', 'glm-4-flash'),
    'max_tokens': 8192,
    'temperature': 0.3,
}


# ---------------------------------------------------------------------------
# SDK MCP Tools
# ---------------------------------------------------------------------------

def _create_tools():
    """Create SDK MCP tools for the East Money research workflow."""
    from claude_agent_sdk import tool, create_sdk_mcp_server

    @tool(
        name="fetch_research_reports",
        description=(
            "从东方财富获取行业研报元数据和正文文本。"
            "支持个股研报和行业研报两种搜索模式，自动合并去重。"
            "返回研报列表和合并后的文本（用于LLM输入）。"
        ),
        input_schema={
            "stock_code": str,
            "industry_name": str,
            "sw_l1_name": str,
            "max_reports": int,
        },
    )
    async def fetch_reports(args):
        stock_code = args["stock_code"]
        industry_name = args["industry_name"]
        sw_l1_name = args.get("sw_l1_name", "")
        max_reports = args.get("max_reports", 8)

        config = {**DEFAULT_FETCHER_CONFIG, 'max_reports': max_reports}
        fetcher = ReportFetcher(config)
        result = fetcher.fetch_reports_for_chapter(
            stock_code=stock_code,
            industry_name=industry_name,
            sw_l1_name=sw_l1_name or industry_name,
        )

        report_count = len(result.get('reports', []))
        text_len = len(result.get('report_text', ''))
        errors = len(result.get('download_errors', []))

        summary = json.dumps({
            'report_count': report_count,
            'text_length': text_len,
            'download_errors': errors,
            'reports': result.get('reports', []),
            'report_text_preview': result.get('report_text', '')[:3000],
        }, ensure_ascii=False)

        return {
            "content": [{"type": "text", "text": summary}],
            # Store full data in a temporary side channel
            "_full_result": result,
        }

    @tool(
        name="generate_industry_analysis",
        description=(
            "调用LLM基于研报文本生成结构化行业分析。"
            "输出包含行业概述、产业链、市场规模、竞争格局等章节，"
            "附带脚注引用和市场规模数据。"
        ),
        input_schema={
            "company_name": str,
            "stock_code": str,
            "industry_name": str,
            "sw_l1_name": str,
            "report_text": str,
            "peer_summary": str,
        },
    )
    async def generate_analysis(args):
        llm_config = {
            **DEFAULT_LLM_CONFIG,
            'api_key': os.environ.get('LLM_API_KEY', '') or os.environ.get('ZHIPUAI_API_KEY', ''),
        }
        writer = LLMWriter(llm_config)
        result = writer.generate_industry_analysis(
            company_name=args["company_name"],
            stock_code=args["stock_code"],
            industry_name=args["industry_name"],
            sw_l1_name=args.get("sw_l1_name", args["industry_name"]),
            report_text=args["report_text"],
            peer_data_summary=args.get("peer_summary", ""),
        )

        if result:
            return {
                "content": [{"type": "text", "text": json.dumps(result, ensure_ascii=False)}],
            }
        else:
            return {
                "content": [{"type": "text", "text": json.dumps({"error": "LLM analysis failed"}, ensure_ascii=False)}],
                "is_error": True,
            }

    return create_sdk_mcp_server(
        name="eastmoney-research",
        version="1.0.0",
        tools=[fetch_reports, generate_analysis],
    )


# ---------------------------------------------------------------------------
# Agent Class
# ---------------------------------------------------------------------------

class EastMoneyResearchAgent:
    """East Money Research Report Agent.

    Uses Claude Agent SDK to orchestrate:
    1. Fetch research reports from East Money
    2. Extract text from report web pages
    3. Generate structured analysis via LLM

    Falls back to direct pipeline if SDK is unavailable.
    """

    def __init__(self, config: Optional[dict] = None):
        self.config = config or {}
        self.fetcher_config = {**DEFAULT_FETCHER_CONFIG, **self.config.get('report_fetcher', {})}
        self.llm_config = {**DEFAULT_LLM_CONFIG, **self.config.get('llm', {})}

    async def run_sdk_agent(
        self,
        stock_code: str,
        company_name: str,
        industry_name: str,
        sw_l1_name: str = "",
    ) -> dict:
        """Run using Claude Agent SDK for intelligent orchestration."""
        from claude_agent_sdk import query, ClaudeAgentOptions

        mcp_server = _create_tools()

        prompt = (
            f"请为 {company_name} ({stock_code}) 进行行业研究分析。\n\n"
            f"行业: {industry_name}\n"
            f"SW分类: {sw_l1_name or industry_name}\n\n"
            "请按以下步骤操作:\n"
            "1. 使用 fetch_research_reports 工具获取行业研报\n"
            "2. 如果获取到研报文本，使用 generate_industry_analysis 工具生成结构化分析\n"
            "3. 总结完整的分析结果\n"
        )

        options = ClaudeAgentOptions(
            system_prompt=(
                "你是一个专业的金融行业研究分析师。你的任务是:\n"
                "1. 从东方财富获取相关行业研报\n"
                "2. 基于研报内容生成结构化的行业分析\n"
                "3. 确保所有数据有来源引用，不编造信息\n"
                "请使用中文进行分析和输出。"
            ),
            mcp_servers={"eastmoney-research": mcp_server},
            permission_mode="bypassPermissions",
            max_turns=20,
            max_budget_usd=0.5,
        )

        result_text_parts = []
        async for event in query(prompt=prompt, options=options):
            from claude_agent_sdk import ResultMessage, AssistantMessage

            if isinstance(event, ResultMessage):
                for block in event.content:
                    if hasattr(block, 'text'):
                        result_text_parts.append(block.text)

        full_text = '\n'.join(result_text_parts)
        return {
            'mode': 'sdk_agent',
            'stock_code': stock_code,
            'company_name': company_name,
            'analysis': full_text,
        }

    def run_pipeline(
        self,
        stock_code: str,
        company_name: str,
        industry_name: str,
        sw_l1_name: str = "",
    ) -> dict:
        """Run the direct pipeline without SDK (synchronous)."""
        print(f"[1/2] 获取研报数据: {company_name} ({stock_code})...")

        fetcher = ReportFetcher(self.fetcher_config)
        report_result = fetcher.fetch_reports_for_chapter(
            stock_code=stock_code,
            industry_name=industry_name,
            sw_l1_name=sw_l1_name or industry_name,
        )

        reports = report_result.get('reports', [])
        report_text = report_result.get('report_text', '')
        errors = report_result.get('download_errors', [])

        print(f"  找到 {len(reports)} 篇研报, 提取文本 {len(report_text)} 字")
        if errors:
            print(f"  {len(errors)} 篇提取失败")

        analysis = None
        if report_text:
            print(f"[2/2] 调用LLM生成行业分析...")
            llm_config = {
                **self.llm_config,
                'api_key': (self.llm_config.get('api_key', '')
                            or os.environ.get('LLM_API_KEY', '')
                            or os.environ.get('ZHIPUAI_API_KEY', '')),
            }
            writer = LLMWriter(llm_config)
            analysis = writer.generate_industry_analysis(
                company_name=company_name,
                stock_code=stock_code,
                industry_name=industry_name,
                sw_l1_name=sw_l1_name or industry_name,
                report_text=report_text,
                peer_data_summary='',
            )
            if analysis:
                print(f"  LLM分析完成, 使用模型: {analysis.get('model_used', 'unknown')}")
            else:
                print("  LLM分析失败")
        else:
            print("[2/2] 无研报文本, 跳过LLM分析")

        return {
            'mode': 'pipeline',
            'stock_code': stock_code,
            'company_name': company_name,
            'reports': reports,
            'report_text': report_text,
            'analysis': analysis,
            'errors': errors,
        }

    async def run(
        self,
        stock_code: str,
        company_name: str,
        industry_name: str,
        sw_l1_name: str = "",
        use_sdk: bool = True,
    ) -> dict:
        """Run the agent, trying SDK first, falling back to pipeline."""
        if use_sdk:
            try:
                from claude_agent_sdk import query  # noqa: F401
                return await self.run_sdk_agent(stock_code, company_name, industry_name, sw_l1_name)
            except ImportError:
                print("claude-agent-sdk 未安装, 使用直接管线模式")
            except Exception as e:
                print(f"SDK agent 运行失败: {e}, 回退到直接管线模式")

        return self.run_pipeline(stock_code, company_name, industry_name, sw_l1_name)


# ---------------------------------------------------------------------------
# Convenience function
# ---------------------------------------------------------------------------

async def run_research(
    stock_code: str,
    company_name: str,
    industry_name: str = "",
    sw_l1_name: str = "",
    use_sdk: bool = True,
    config: Optional[dict] = None,
) -> dict:
    """Quick entry point for running the East Money research agent.

    Args:
        stock_code: Stock code like '002276.SZ'
        company_name: Company name in Chinese
        industry_name: SW industry name (L3 preferred)
        sw_l1_name: Comma-separated SW L1,L2,L3 names for keyword expansion
        use_sdk: Try Claude Agent SDK first (falls back to pipeline)
        config: Optional config dict for fetcher and LLM

    Returns:
        Dict with reports, analysis, and metadata
    """
    agent = EastMoneyResearchAgent(config=config)
    return await agent.run(
        stock_code=stock_code,
        company_name=company_name,
        industry_name=industry_name or company_name,
        sw_l1_name=sw_l1_name,
        use_sdk=use_sdk,
    )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description='东方财富行业研报Agent - 获取研报并生成结构化行业分析',
    )
    parser.add_argument('--stock', required=True, help='股票代码 (如 002276.SZ)')
    parser.add_argument('--name', required=True, help='公司名称')
    parser.add_argument('--industry', default='', help='申万行业名称 (L3)')
    parser.add_argument('--sw-l1', default='', help='SW L1,L2,L3 层级名称 (逗号分隔)')
    parser.add_argument('--output', default='', help='输出 JSON 文件路径')
    parser.add_argument('--no-sdk', action='store_true', help='跳过 SDK, 使用直接管线模式')
    parser.add_argument('--max-reports', type=int, default=8, help='最大研报数')

    args = parser.parse_args()

    config = {
        'report_fetcher': {'max_reports': args.max_reports},
    }

    agent = EastMoneyResearchAgent(config=config)

    if args.no_sdk:
        result = agent.run_pipeline(
            stock_code=args.stock,
            company_name=args.name,
            industry_name=args.industry or args.name,
            sw_l1_name=args.sw_l1,
        )
    else:
        result = asyncio.run(agent.run(
            stock_code=args.stock,
            company_name=args.name,
            industry_name=args.industry or args.name,
            sw_l1_name=args.sw_l1,
            use_sdk=True,
        ))

    output = json.dumps(result, ensure_ascii=False, indent=2, default=str)

    if args.output:
        os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output)
        print(f"\n结果已保存到 {args.output}")
    else:
        print(f"\n{output}")


if __name__ == '__main__':
    main()
