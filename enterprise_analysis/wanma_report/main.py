#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增项目企业基本面风险分析报告 - 主程序
从Tushare获取数据生成企业财务风险评估报告
"""

import os
import sys
import yaml
import pandas as pd
from datetime import datetime
from typing import Dict, Any, Optional

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tushare as ts
from utils.tushare_client import TushareClient
from utils.financial_scoring import FinancialScoringEngine
from utils.industry_analyzer import IndustryAnalyzer


class DingZengRiskReportGenerator:
    """定增项目企业基本面风险分析报告生成器"""

    def __init__(self, config_path: str = None):
        """初始化报告生成器

        Args:
            config_path: 配置文件路径
        """
        if config_path is None:
            config_path = os.path.join(os.path.dirname(__file__), 'config.yaml')

        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)

        # 获取Tushare Token (优先使用环境变量)
        tushare_token = self.config['tushare'].get('token', '') or os.environ.get('TUSHARE_TOKEN', '')
        if not tushare_token:
            raise ValueError("请设置TUSHARE_TOKEN环境变量或在config.yaml中配置token")

        # 初始化Tushare客户端
        self.ts_client = TushareClient(tushare_token)
        self.pro = self.ts_client.get_pro()

        # 初始化分析器
        self.scoring_engine = FinancialScoringEngine(self.config['scoring'])
        self.industry_analyzer = IndustryAnalyzer(self.pro, self.config['industry'])

        # 项目信息
        self.project_info = self.config['project']

    def load_all_data(self) -> Dict[str, Any]:
        """加载所有需要的数据

        Returns:
            包含所有分析数据的字典
        """
        print("=" * 60)
        print("开始加载数据...")
        print("=" * 60)

        data = {}

        # 1. 获取公司基本信息
        print("\n[1/6] 获取公司基本信息...")
        data['company_info'] = self._get_company_info()

        # 2. 获取财务指标数据
        print("[2/6] 获取财务指标数据...")
        data['financial_indicators'] = self._get_financial_indicators()

        # 3. 获取财务报表数据
        print("[3/6] 获取财务报表数据...")
        data['financial_statements'] = self._get_financial_statements()

        # 4. 获取行业对比数据
        print("[4/6] 获取行业对比数据...")
        data['industry_data'] = self._get_industry_comparison()

        # 5. 计算财务评分
        print("[5/6] 计算财务评分...")
        data['scoring_result'] = self._calculate_financial_score()

        # 6. 获取估值数据
        print("[6/6] 获取估值数据...")
        data['valuation_data'] = self._get_valuation_data()

        print("\n数据加载完成!")
        return data

    def _get_company_info(self) -> Dict[str, Any]:
        """获取公司基本信息"""
        stock_code = self.project_info['stock_code']

        # 获取公司基本信息
        company_info = self.pro.stock_basic(ts_code=stock_code, fields='ts_code,name,industry,market,list_date')

        if company_info is None or len(company_info) == 0:
            print(f"警告: 无法获取股票 {stock_code} 的基本信息")
            return {}

        info = company_info.iloc[0].to_dict()

        # 获取公司详细信息
        try:
            name_en = self.pro.namechange(ts_code=stock_code, fields='name_en,short_name')
            if name_en is not None and len(name_en) > 0:
                info['name_en'] = name_en.iloc[0].get('name_en', '')
                info['short_name'] = name_en.iloc[0].get('short_name', '')
        except Exception as e:
            print(f"获取公司英文名称时出错: {e}")

        return info

    def _get_financial_indicators(self) -> pd.DataFrame:
        """获取财务指标数据"""
        stock_code = self.project_info['stock_code']

        # 获取近5年的财务指标
        end_date = datetime.now().strftime('%Y%m%d')
        start_year = int(datetime.now().strftime('%Y')) - 5
        start_date = f"{start_year}0101"

        df = self.pro.fina_indicator(
            ts_code=stock_code,
            start_date=start_date,
            end_date=end_date
        )

        # 只保留年报数据
        if df is not None and len(df) > 0:
            df = df[df['end_date'].str.contains('1231', na=False)]

        return df if df is not None else pd.DataFrame()

    def _get_financial_statements(self) -> Dict[str, pd.DataFrame]:
        """获取财务报表数据"""
        stock_code = self.project_info['stock_code']

        end_date = datetime.now().strftime('%Y%m%d')
        start_year = int(datetime.now().strftime('%Y')) - 5
        start_date = f"{start_year}0101"

        statements = {}

        # 资产负债表
        statements['balance_sheet'] = self.pro.balancesheet(
            ts_code=stock_code,
            start_date=start_date,
            end_date=end_date
        )

        # 利润表
        statements['income_statement'] = self.pro.income(
            ts_code=stock_code,
            start_date=start_date,
            end_date=end_date
        )

        # 现金流量表
        statements['cashflow'] = self.pro.cashflow(
            ts_code=stock_code,
            start_date=start_date,
            end_date=end_date
        )

        return statements

    def _get_industry_comparison(self) -> pd.DataFrame:
        """获取行业对比数据"""
        industry_codes = self.config['industry']['codes']
        end_date = datetime.now().strftime('%Y%m%d')
        start_year = int(datetime.now().strftime('%Y')) - 5
        start_date = f"{start_year}0101"

        # 获取行业内所有公司的财务指标
        df_list = []
        for code in industry_codes:
            try:
                df = self.pro.fina_indicator(
                    ts_code=code,
                    start_date=start_date,
                    end_date=end_date
                )
                if df is not None and len(df) > 0:
                    df = df[df['end_date'].str.contains('1231', na=False)]
                    df_list.append(df)
            except Exception as e:
                print(f"获取 {code} 财务指标时出错: {e}")

        if df_list:
            return pd.concat(df_list, ignore_index=True)
        return pd.DataFrame()

    def _calculate_financial_score(self) -> Dict[str, Any]:
        """计算财务评分"""
        financial_data = self._get_financial_indicators()

        if financial_data.empty:
            return {
                'total_score': 0,
                'grade': '暂无数据',
                'dimensions': {}
            }

        # 使用评分引擎计算评分
        result = self.scoring_engine.calculate_score(financial_data)

        return result

    def _get_valuation_data(self) -> Dict[str, Any]:
        """获取估值数据"""
        stock_code = self.project_info['stock_code']

        data = {}

        # 获取每日行情数据
        end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.now() - pd.Timedelta(days=365)).strftime('%Y%m%d')

        df_daily = self.pro.daily(ts_code=stock_code, start_date=start_date, end_date=end_date)

        if df_daily is not None and len(df_daily) > 0:
            data['current_price'] = df_daily.iloc[-1]['close']
            data['pe_ratio'] = df_daily.iloc[-1].get('pe', None)
            data['pb_ratio'] = df_daily.iloc[-1].get('pb', None)

        # 获取估值指标
        try:
            df_val = self.pro.daily_basic(ts_code=stock_code, start_date=start_date, end_date=end_date)
            if df_val is not None and len(df_val) > 0:
                latest = df_val.iloc[-1]
                data['market_cap'] = latest.get('total_mv', None)  # 总市值(万元)
                data['float_market_cap'] = latest.get('float_mv', None)  # 流通市值(万元)
        except Exception as e:
            print(f"获取估值数据时出错: {e}")

        return data

    def generate_report(self, output_path: str = None) -> str:
        """生成完整的风险评估报告

        Args:
            output_path: 输出文件路径

        Returns:
            生成的报告文件路径
        """
        print("\n" + "=" * 60)
        print("开始生成报告...")
        print("=" * 60)

        # 加载所有数据
        all_data = self.load_all_data()

        # 生成各章节内容
        chapters = []

        # 第一章：项目概况
        print("\n生成第一章：项目概况...")
        from chapters.chapter01_overview import Chapter01Overview
        ch01 = Chapter01Overview(self.config, all_data)
        chapters.append(ch01.generate())

        # 第二章：风险概述
        print("生成第二章：风险概述...")
        from chapters.chapter02_risk_summary import Chapter02RiskSummary
        ch02 = Chapter02RiskSummary(self.config, all_data)
        chapters.append(ch02.generate())

        # 第三章：财务风险分析
        print("生成第三章：财务风险分析...")
        from chapters.chapter03_financial import Chapter03Financial
        ch03 = Chapter03Financial(self.config, all_data, self.scoring_engine)
        chapters.append(ch03.generate())

        # 第四章：估值风险
        print("生成第四章：估值风险...")
        from chapters.chapter04_valuation import Chapter04Valuation
        ch04 = Chapter04Valuation(self.config, all_data)
        chapters.append(ch04.generate())

        # 第五章：退出风险
        print("生成第五章：退出风险...")
        from chapters.chapter05_exit import Chapter05Exit
        ch05 = Chapter05Exit(self.config, all_data)
        chapters.append(ch05.generate())

        # 第六章：招引落地
        print("生成第六章：招引落地...")
        from chapters.chapter06_investment import Chapter06Investment
        ch06 = Chapter06Investment(self.config, all_data)
        chapters.append(ch06.generate())

        # 生成Word文档
        print("\n生成Word文档...")
        if output_path is None:
            output_dir = os.path.join(os.path.dirname(__file__), self.config['report']['output_dir'])
            os.makedirs(output_dir, exist_ok=True)
            output_path = os.path.join(output_dir, f"{self.project_info['name']}定增项目企业基本面风险分析报告.docx")

        self._generate_word_document(chapters, output_path)

        print(f"\n报告生成完成: {output_path}")
        return output_path

    def _generate_word_document(self, chapters: list, output_path: str) -> None:
        """生成Word文档

        Args:
            chapters: 章节内容列表
            output_path: 输出文件路径
        """
        from docx import Document
        from docx.shared import Pt, Inches
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        doc = Document()

        # 设置文档标题
        title = doc.add_heading(f"{self.project_info['name']}定增项目企业基本面风险分析报告", level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 添加报告信息
        info_para = doc.add_paragraph()
        info_para.add_run(f"报告日期: {self.project_info['report_date']}").bold = True
        info_para.add_run(f"\n版本: {self.project_info['version']}")
        info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_page_break()

        # 添加各章节内容
        for chapter in chapters:
            for element in chapter:
                if element['type'] == 'heading':
                    doc.add_heading(element['content'], level=element.get('level', 1))
                elif element['type'] == 'paragraph':
                    para = doc.add_paragraph(element['content'])
                elif element['type'] == 'table':
                    self._add_table_to_doc(doc, element['data'], element.get('headers', []))

        # 保存文档
        doc.save(output_path)

    def _add_table_to_doc(self, doc, data: list, headers: list) -> None:
        """添加表格到Word文档

        Args:
            doc: Word文档对象
            data: 表格数据
            headers: 表头
        """
        if not data:
            return

        table = doc.add_table(rows=1 + len(data), cols=len(headers) if headers else len(data[0]))
        table.style = 'Light Grid Accent 1'

        # 添加表头
        header_cells = table.rows[0].cells
        for i, header in enumerate(headers if headers else data[0].keys()):
            header_cells[i].text = str(header)

        # 添加数据行
        for row_idx, row_data in enumerate(data):
            row_cells = table.rows[row_idx + 1].cells
            if isinstance(row_data, dict):
                for col_idx, value in enumerate(row_data.values()):
                    row_cells[col_idx].text = str(value) if value is not None else ''
            else:
                for col_idx, value in enumerate(row_data):
                    row_cells[col_idx].text = str(value) if value is not None else ''


def main():
    """主函数"""
    print("=" * 60)
    print("定增项目企业基本面风险分析报告生成程序")
    print("=" * 60)

    # 创建报告生成器
    generator = DingZengRiskReportGenerator()

    # 生成报告
    output_path = generator.generate_report()

    print(f"\n报告已生成: {output_path}")


if __name__ == "__main__":
    main()
