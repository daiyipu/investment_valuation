#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
行业分析器
用于行业数据对比和阈值计算
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, List, Optional


class IndustryAnalyzer:
    """行业分析器

    用于:
    - 获取行业内公司的财务指标
    - 计算行业阈值和分位数
    - 进行同业对比分析
    """

    def __init__(self, pro_api, config: Dict[str, Any]):
        """初始化行业分析器

        Args:
            pro_api: Tushare pro API对象
            config: 行业配置
        """
        self.pro = pro_api
        self.config = config
        self.industry_name = config.get('industry', {}).get('name', '')
        self.industry_codes = config.get('industry', {}).get('codes', [])

        # 缓存行业数据
        self._industry_data_cache = None

    def get_industry_data(self, refresh: bool = False) -> pd.DataFrame:
        """获取行业数据

        Args:
            refresh: 是否强制刷新缓存

        Returns:
            行业公司财务指标数据
        """
        if self._industry_data_cache is not None and not refresh:
            return self._industry_data_cache

        df_list = []
        for code in self.industry_codes:
            try:
                df = self.pro.fina_indicator(ts_code=code)
                if df is not None and len(df) > 0:
                    df = df[df['end_date'].str.contains('1231', na=False)]
                    df['ts_code'] = code
                    df_list.append(df)
            except Exception as e:
                print(f"获取 {code} 财务指标失败: {e}")

        if df_list:
            self._industry_data_cache = pd.concat(df_list, ignore_index=True)
            return self._industry_data_cache

        return pd.DataFrame()

    def calculate_industry_thresholds(
        self,
        field: str,
        percentiles: List[float] = None
    ) -> Dict[str, float]:
        """计算行业指标阈值

        Args:
            field: 指标字段名
            percentiles: 分位数列表，默认[10, 25, 50, 75, 90]

        Returns:
            各分位数阈值字典
        """
        if percentiles is None:
            percentiles = [10, 25, 50, 75, 90]

        industry_data = self.get_industry_data()
        if industry_data.empty or field not in industry_data.columns:
            return {}

        # 获取所有有效数据
        values = industry_data[field].dropna().values
        if len(values) == 0:
            return {}

        # 计算各分位数
        thresholds = {}
        for p in percentiles:
            thresholds[f'p{p}'] = np.percentile(values, p)

        thresholds['mean'] = np.mean(values)
        thresholds['std'] = np.std(values)
        thresholds['min'] = np.min(values)
        thresholds['max'] = np.max(values)

        return thresholds

    def compare_with_industry(
        self,
        company_data: pd.DataFrame,
        fields: List[str] = None
    ) -> pd.DataFrame:
        """与行业平均对比

        Args:
            company_data: 公司财务指标数据
            fields: 要对比的字段列表

        Returns:
            对比结果DataFrame
        """
        if fields is None:
            fields = [
                'roe', 'gross_margin', 'net_margin', 'current_ratio',
                'quick_ratio', 'debt_to_assets', 'inv_turn', 'arturn_days',
                'net_profit_yoy', 'revenue_yoy'
            ]

        industry_data = self.get_industry_data()
        if industry_data.empty:
            return pd.DataFrame()

        results = []
        for field in fields:
            if field not in company_data.columns:
                continue

            # 公司值
            company_value = self._get_latest_value(company_data, field)

            # 行业阈值
            thresholds = self.calculate_industry_thresholds(field)

            # 计算公司百分位
            company_percentile = None
            if thresholds and company_value is not None:
                p_values = [thresholds.get(f'p{p}') for p in [10, 25, 50, 75, 90]]
                p_values = [v for v in p_values if v is not None]
                if p_values:
                    company_percentile = sum(1 for v in p_values if v < company_value) / len(p_values) * 100

            results.append({
                '指标': field,
                '公司值': company_value,
                '行业P25': thresholds.get('p25'),
                '行业P50': thresholds.get('p50'),
                '行业P75': thresholds.get('p75'),
                '公司百分位': company_percentile
            })

        return pd.DataFrame(results)

    def get_peer_companies(
        self,
        company_data: pd.DataFrame,
        year: str = None
    ) -> pd.DataFrame:
        """获取同业公司列表

        Args:
            company_data: 公司财务指标数据
            year: 选定的年份

        Returns:
            同业公司数据
        """
        industry_data = self.get_industry_data()

        if industry_data.empty:
            return pd.DataFrame()

        # 如果指定年份，筛选该年份数据
        if year:
            industry_data = industry_data[industry_data['end_date'].str.startswith(year)]

        return industry_data

    def calculate_percentile(
        self,
        value: float,
        field: str,
        year: str = None
    ) -> float:
        """计算单个公司在行业中的百分位

        Args:
            value: 公司指标值
            field: 指标字段名
            year: 选定的年份

        Returns:
            百分位(0-100)
        """
        thresholds = self.calculate_industry_thresholds(field)
        if not thresholds:
            return 50  # 默认中间值

        p_values = [thresholds.get(f'p{p}') for p in [10, 25, 50, 75, 90]]
        p_values = [v for v in p_values if v is not None]

        if not p_values:
            return 50

        percentile = sum(1 for v in p_values if v < value) / len(p_values) * 100
        return max(0, min(100, percentile))

    def _get_latest_value(self, df: pd.DataFrame, field: str) -> Optional[float]:
        """获取最新一期有效数据

        Args:
            df: 财务数据
            field: 字段名

        Returns:
            最新有效值
        """
        if field not in df.columns:
            return None

        df_sorted = df.sort_values('end_date', ascending=False)

        for idx, row in df_sorted.iterrows():
            value = row.get(field)
            if value is not None and not pd.isna(value):
                return float(value)

        return None

    def get_industry_summary(self) -> Dict[str, Any]:
        """获取行业汇总信息

        Returns:
            行业汇总数据
        """
        industry_data = self.get_industry_data()
        if industry_data.empty:
            return {}

        summary = {
            '行业名称': self.industry_name,
            '公司数量': len(self.industry_codes),
            '数据期数': len(industry_data['end_date'].unique()) if 'end_date' in industry_data.columns else 0
        }

        # 计算主要指标的行业中位数
        key_fields = ['roe', 'gross_margin', 'net_margin', 'current_ratio', 'debt_to_assets']
        for field in key_fields:
            if field in industry_data.columns:
                median = industry_data[field].median()
                summary[f'{field}_中位数'] = round(median, 4) if median is not None else None

        return summary

    def export_industry_report(self, output_path: str = None) -> None:
        """导出行业分析报告

        Args:
            output_path: 输出文件路径
        """
        industry_data = self.get_industry_data()
        if industry_data.empty:
            print("没有行业数据可导出")
            return

        if output_path is None:
            output_path = f"{self.industry_name}_行业分析.xlsx"

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 导出原始数据
            industry_data.to_excel(writer, sheet_name='原始数据', index=False)

            # 导出汇总信息
            summary = self.get_industry_summary()
            summary_df = pd.DataFrame([summary])
            summary_df.to_excel(writer, sheet_name='行业汇总', index=False)

            # 导出各指标阈值
            key_fields = ['roe', 'gross_margin', 'net_margin', 'current_ratio', 'debt_to_assets']
            threshold_rows = []
            for field in key_fields:
                thresholds = self.calculate_industry_thresholds(field)
                if thresholds:
                    thresholds['指标'] = field
                    threshold_rows.append(thresholds)

            if threshold_rows:
                threshold_df = pd.DataFrame(threshold_rows)
                threshold_df.to_excel(writer, sheet_name='指标阈值', index=False)

        print(f"行业分析报告已导出: {output_path}")
