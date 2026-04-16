#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务评分引擎
基于EFAES评分系统的企业财务评分计算
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Optional, List


class FinancialScoringEngine:
    """财务评分引擎

    基于EFAES评分系统的四维度财务评分:
    - 运营能力 (30%权重)
    - 盈利能力 (30%权重)
    - 偿债能力 (20%权重)
    - 成长能力 (20%权重)
    """

    # 评分等级定义
    GRADE_LEVELS = {
        '优秀': (90, 100),
        '良好': (80, 89),
        '中等': (70, 79),
        '较差': (60, 69),
        '极差': (0, 59)
    }

    def __init__(self, config: Dict[str, Any]):
        """初始化评分引擎

        Args:
            config: 评分配置，包含各维度和指标定义
        """
        self.config = config
        # 配置可能是完整的config或嵌套的scoring部分
        self.scoring_config = config.get('scoring', config)
        self._validate_config()

    def _validate_config(self) -> None:
        """验证配置完整性"""
        required_dimensions = ['operations', 'profitability', 'solvency', 'growth']
        for dim in required_dimensions:
            if dim not in self.scoring_config:
                raise ValueError(f"配置缺少必要维度: {dim}")

        # 验证权重总和
        total_weight = sum(self.scoring_config[dim].get('weight', 0) for dim in required_dimensions)
        if abs(total_weight - 1.0) > 0.01:
            print(f"警告: 权重总和为 {total_weight}，期望为 1.0")

    def calculate_score(self, financial_data: pd.DataFrame) -> Dict[str, Any]:
        """计算财务评分

        Args:
            financial_data: 财务指标数据

        Returns:
            评分结果，包含总分、各维度得分和详细指标评分
        """
        if financial_data.empty:
            return self._empty_result()

        result = {
            'total_score': 0,
            'grade': '暂无数据',
            'dimensions': {},
            'indicators': {},
            'industry_percentile': {}
        }

        dimension_scores = {}
        dimension_details = {}

        # 计算各维度评分
        for dimension, config in self.scoring_config.items():
            dim_score, dim_indicators = self._calculate_dimension_score(
                dimension, config, financial_data
            )
            dimension_scores[dimension] = dim_score
            dimension_details[dimension] = dim_indicators
            result['dimensions'][dimension] = {
                'score': dim_score,
                'weight': config.get('weight', 0),
                'weighted_score': dim_score * config.get('weight', 0)
            }
            result['indicators'][dimension] = dim_indicators

        # 计算总分
        total_score = sum(
            result['dimensions'][dim]['weighted_score']
            for dim in result['dimensions']
        )
        result['total_score'] = round(total_score, 2)

        # 确定评分等级
        result['grade'] = self._get_grade(result['total_score'])

        # 计算行业百分位
        result['industry_percentile'] = self._calculate_industry_percentile(financial_data)

        return result

    def _calculate_dimension_score(
        self,
        dimension: str,
        config: Dict[str, Any],
        financial_data: pd.DataFrame
    ) -> tuple:
        """计算单个维度评分

        Args:
            dimension: 维度名称
            config: 维度配置
            financial_data: 财务数据

        Returns:
            (维度得分, 指标详情列表)
        """
        indicators = config.get('indicators', [])
        if not indicators:
            return 0, []

        indicator_scores = []
        valid_count = 0

        for ind_config in indicators:
            field = ind_config.get('field')
            ind_name = ind_config.get('name')
            reverse = ind_config.get('reverse', False)

            if field not in financial_data.columns:
                indicator_scores.append({
                    'name': ind_name,
                    'field': field,
                    'value': None,
                    'score': None,
                    'percentile': None,
                    'status': '数据缺失'
                })
                continue

            # 获取最新一期有效数据
            value = self._get_latest_value(financial_data, field)

            if value is None or pd.isna(value):
                indicator_scores.append({
                    'name': ind_name,
                    'field': field,
                    'value': None,
                    'score': None,
                    'percentile': None,
                    'status': '数据缺失'
                })
                continue

            # 计算单项评分
            score = self._calculate_indicator_score(value, field, reverse)
            indicator_scores.append({
                'name': ind_name,
                'field': field,
                'value': round(value, 4) if value is not None else None,
                'score': round(score, 2) if score is not None else None,
                'status': '正常'
            })
            valid_count += 1

        # 计算维度平均分
        valid_scores = [s['score'] for s in indicator_scores if s['score'] is not None]
        if valid_scores:
            dimension_score = np.mean(valid_scores)
        else:
            dimension_score = 0

        return dimension_score, indicator_scores

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

        # 按日期排序，获取最新一期
        df_sorted = df.sort_values('end_date', ascending=False)

        for idx, row in df_sorted.iterrows():
            value = row.get(field)
            if value is not None and not pd.isna(value):
                return float(value)

        return None

    def _calculate_indicator_score(
        self,
        value: float,
        field: str,
        reverse: bool = False
    ) -> float:
        """计算单项指标评分

        使用行业阈值进行百分位计算，然后转换为百分制评分

        Args:
            value: 指标值
            field: 字段名
            reverse: 是否反向(越小越好)

        Returns:
            指标评分(0-100)
        """
        # 获取行业阈值
        thresholds = self._get_industry_thresholds(field)

        if not thresholds:
            # 如果没有行业阈值，使用简单线性评分
            return self._linear_score(value, reverse)

        # 计算百分位
        sorted_values = sorted(thresholds)
        percentile = sum(1 for v in sorted_values if v < value) / len(sorted_values) * 100

        if reverse:
            percentile = 100 - percentile

        # 限制在0-100范围内
        score = max(0, min(100, percentile))

        return score

    def _get_industry_thresholds(self, field: str) -> List[float]:
        """获取行业阈值

        Args:
            field: 字段名

        Returns:
            行业阈值列表
        """
        # TODO: 从行业数据库或配置中获取实际阈值
        # 这里使用简化的示例阈值
        default_thresholds = {
            'roe': [5, 8, 12, 15, 20],
            'gross_margin': [15, 20, 25, 35, 45],
            'current_ratio': [0.8, 1.0, 1.5, 2.0, 3.0],
            'debt_to_assets': [70, 60, 50, 40, 30],
            'net_margin': [2, 5, 8, 12, 15],
        }

        return default_thresholds.get(field, [])

    def _linear_score(self, value: float, reverse: bool = False) -> float:
        """使用线性评分

        Args:
            value: 指标值
            reverse: 是否反向

        Returns:
            评分
        """
        # 简化的线性评分逻辑
        # 根据指标类型设置合理的范围
        if reverse:
            # 反向指标：值越小越好，假设合理范围是0-100
            score = max(0, 100 - value)
        else:
            # 正向指标：值越大越好，假设合理范围是0-100
            score = min(100, value)

        return score

    def _calculate_industry_percentile(self, financial_data: pd.DataFrame) -> Dict[str, float]:
        """计算行业百分位

        Args:
            financial_data: 财务数据

        Returns:
            各指标的百分位
        """
        percentile = {}

        for dimension, config in self.scoring_config.items():
            for ind_config in config.get('indicators', []):
                field = ind_config.get('field')
                if field in financial_data.columns:
                    value = self._get_latest_value(financial_data, field)
                    if value is not None:
                        percentile[field] = self._calculate_percentile(value, field)

        return percentile

    def _calculate_percentile(self, value: float, field: str) -> float:
        """计算单个指标的百分位

        Args:
            value: 指标值
            field: 字段名

        Returns:
            百分位(0-100)
        """
        thresholds = self._get_industry_thresholds(field)
        if not thresholds:
            return 50  # 默认中间值

        sorted_values = sorted(thresholds)
        percentile = sum(1 for v in sorted_values if v < value) / len(sorted_values) * 100
        return round(percentile, 2)

    def _get_grade(self, score: float) -> str:
        """根据评分获取等级

        Args:
            score: 评分

        Returns:
            等级名称
        """
        for grade, (low, high) in self.GRADE_LEVELS.items():
            if low <= score <= high:
                return grade
        return '极差'

    def _empty_result(self) -> Dict[str, Any]:
        """返回空结果"""
        return {
            'total_score': 0,
            'grade': '暂无数据',
            'dimensions': {},
            'indicators': {},
            'industry_percentile': {}
        }

    def get_dimension_summary(self, result: Dict[str, Any]) -> pd.DataFrame:
        """获取维度评分汇总表

        Args:
            result: 评分结果

        Returns:
            维度评分汇总DataFrame
        """
        rows = []
        for dim, data in result.get('dimensions', {}).items():
            rows.append({
                '维度': self._get_dimension_name(dim),
                '权重': f"{data.get('weight', 0) * 100:.0f}%",
                '得分': data.get('score', 0),
                '加权得分': round(data.get('weighted_score', 0), 2)
            })

        df = pd.DataFrame(rows)
        if not df.empty:
            df.loc['合计'] = ['总分', '100%', result.get('total_score', 0), result.get('total_score', 0)]

        return df

    def _get_dimension_name(self, dim: str) -> str:
        """获取维度中文名称"""
        names = {
            'operations': '运营能力',
            'profitability': '盈利能力',
            'solvency': '偿债能力',
            'growth': '成长能力'
        }
        return names.get(dim, dim)

    def get_indicator_detail_table(self, result: Dict[str, Any]) -> pd.DataFrame:
        """获取指标详情表

        Args:
            result: 评分结果

        Returns:
            指标详情DataFrame
        """
        rows = []

        for dim, indicators in result.get('indicators', {}).items():
            dim_name = self._get_dimension_name(dim)
            config = self.scoring_config.get(dim, {})
            weight = config.get('weight', 0)

            for ind in indicators:
                rows.append({
                    '维度': dim_name,
                    '指标名称': ind.get('name', ''),
                    '指标值': ind.get('value'),
                    '指标得分': ind.get('score'),
                    '状态': ind.get('status', '')
                })

        return pd.DataFrame(rows)

    def get_scoring_summary_text(self, result: Dict[str, Any]) -> str:
        """生成评分总结文本

        Args:
            result: 评分结果

        Returns:
            评分总结文本
        """
        total = result.get('total_score', 0)
        grade = result.get('grade', '暂无数据')

        summary = f"综合财务评分: {total:.2f}分\n"
        summary += f"评级: {grade}\n\n"

        summary += "各维度评分:\n"
        for dim, data in result.get('dimensions', {}).items():
            dim_name = self._get_dimension_name(dim)
            score = data.get('score', 0)
            weight = data.get('weight', 0)
            summary += f"  - {dim_name}(权重{weight*100:.0f}%): {score:.2f}分\n"

        return summary
