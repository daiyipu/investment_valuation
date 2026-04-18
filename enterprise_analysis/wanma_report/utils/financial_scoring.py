#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
财务评分引擎
基于EFAES评分系统的企业财务评分计算（27项指标）

四维度权重：运营能力(15%) + 盈利能力(30%) + 偿债能力(30%) + 成长能力(25%)
偿债能力细分为：短期偿债(15%) + 长期偿债(15%)

评分方法：将指标值与行业分位数阈值(p10/p30/p50/p70/p90)对比，
分为5个等级(优秀/良好/中等/较差/极差)，对应得分比例0.9/0.8/0.7/0.6/0.5
"""

import pandas as pd
import numpy as np
from typing import Dict, Any, Optional, List


class FinancialScoringEngine:
    """财务评分引擎"""

    GRADE_LEVELS = {
        '优秀': (90, 100),
        '良好': (80, 89),
        '中等': (70, 79),
        '较差': (60, 69),
        '极差': (0, 59),
    }

    # 5级评分对应的得分比例
    SCORE_RATIOS = {
        '优秀': 0.9,
        '良好': 0.8,
        '中等': 0.7,
        '较差': 0.6,
        '极差': 0.5,
    }

    # 反向指标（值越小越好）
    REVERSE_INDICATORS = {
        'debt_to_assets',
        'int_to_talcap',
        'debt_to_eqt',
    }

    # 维度中文名映射
    DIMENSION_NAMES = {
        'operations': '运营能力',
        'profitability': '盈利能力',
        'solvency': '偿债能力',
        'growth': '成长能力',
    }

    # 指标中文名映射
    INDICATOR_NAMES = {
        'inv_turn': '存货周转率',
        'ar_turn': '应收账款周转率',
        'ca_turn': '流动资产周转率',
        'assets_turn': '总资产周转率',
        'netprofit_margin': '净利率',
        'grossprofit_margin': '毛利率',
        'roe': '净资产收益率',
        'roe_dt': '扣非净资产收益率',
        'roa': '总资产报酬率',
        'npta': '总资产净利润率',
        'current_ratio': '流动比率',
        'quick_ratio': '速动比率',
        'cash_to_liqdebt': '现金与流动负债比率',
        'cash_to_liqdebt_withinterest': '现金与带息流动负债比率',
        'debt_to_assets': '资产负债率',
        'int_to_talcap': '带息债务/全部投入资本',
        'debt_to_eqt': '产权比率',
        'ebit_to_interest': '利息保障倍数',
        'netprofit_yoy': '净利润同比增长率',
        'dt_netprofit_yoy': '扣非净利润同比增长率',
        'roe_yoy': '净资产收益率同比增长率',
        'tr_yoy': '营业总收入同比增长率',
        'or_yoy': '营业收入同比增长率',
        'equity_yoy': '净资产同比增长率',
        'op_yoy': '营业利润同比增长率',
        'ebt_yoy': '利润总额同比增长率',
        'rd_exp_ratio': '研发费用率',
        # 兼容旧字段名（Tushare fina_indicator可能使用不同字段名）
        'net_margin': '销售净利率',
        'gross_margin': '毛利率',
        'arturn_days': '应收账款周转天数',
        'current_asset_turn': '流动资产周转率',
        'asset_turn': '总资产周转率',
        'net_profit_yoy': '净利润同比增长率',
        'revenue_yoy': '营业收入同比增长率',
    }

    # 字段名映射：将Tushare fina_indicator字段映射到EFAES标准字段
    FIELD_ALIASES = {
        'net_margin': 'netprofit_margin',
        'gross_margin': 'grossprofit_margin',
        'arturn_days': 'ar_turn',
        'current_asset_turn': 'ca_turn',
        'asset_turn': 'assets_turn',
        'net_profit_yoy': 'netprofit_yoy',
        'revenue_yoy': 'or_yoy',
    }

    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.scoring_config = config.get('scoring', config)
        self._industry_data: Optional[pd.DataFrame] = None
        self._financial_statements: Optional[Dict[str, pd.DataFrame]] = None
        self._validate_config()

    def set_industry_data(self, industry_data: pd.DataFrame) -> None:
        self._industry_data = industry_data

    def set_financial_statements(self, statements: Dict[str, pd.DataFrame]) -> None:
        """设置原始财务报表数据，用于计算缺失指标"""
        self._financial_statements = statements

    def _validate_config(self) -> None:
        required_dimensions = ['operations', 'profitability', 'solvency', 'growth']
        for dim in required_dimensions:
            if dim not in self.scoring_config:
                raise ValueError(f"配置缺少必要维度: {dim}")

        total_weight = sum(
            self.scoring_config[dim].get('weight', 0) for dim in required_dimensions
        )
        if abs(total_weight - 1.0) > 0.01:
            print(f"警告: 权重总和为 {total_weight}，期望为 1.0")

    def calculate_score(self, financial_data: pd.DataFrame) -> Dict[str, Any]:
        """计算财务评分"""
        if financial_data.empty:
            return self._empty_result()

        result = {
            'total_score': 0,
            'grade': '暂无数据',
            'dimensions': {},
            'indicators': {},
            'sub_dimensions': {},
        }

        # 计算行业阈值
        thresholds = self._build_industry_thresholds()

        total_score = 0

        for dimension, dim_config in self.scoring_config.items():
            dim_weight = dim_config.get('weight', 0)

            if dimension == 'solvency' and 'sub_dimensions' in dim_config:
                dim_score, dim_indicators, sub_dims = self._calculate_solvency_dimension(
                    dim_config, financial_data, thresholds, dim_weight
                )
                result['sub_dimensions'][dimension] = sub_dims
            else:
                dim_score, dim_indicators = self._calculate_flat_dimension(
                    dimension, dim_config, financial_data, thresholds, dim_weight
                )

            weighted_score = dim_score * dim_weight

            result['dimensions'][dimension] = {
                'score': round(dim_score, 2),
                'weight': dim_weight,
                'weighted_score': round(weighted_score, 2),
            }
            result['indicators'][dimension] = dim_indicators
            total_score += weighted_score

        result['total_score'] = round(total_score, 2)
        result['grade'] = self._get_grade(result['total_score'])
        return result

    def _calculate_flat_dimension(
        self,
        dimension: str,
        dim_config: Dict[str, Any],
        financial_data: pd.DataFrame,
        thresholds: Dict[str, Dict[str, float]],
        dim_weight: float,
    ) -> tuple:
        """计算没有子维度的维度评分（运营、盈利、成长）"""
        indicators = dim_config.get('indicators', [])
        if not indicators:
            return 0, []

        indicator_results = []
        total_indicator_weight = sum(ind.get('weight', 0) for ind in indicators)

        for ind_config in indicators:
            field = ind_config.get('field')
            ind_name = ind_config.get('name', self.INDICATOR_NAMES.get(field, field))
            ind_weight = ind_config.get('weight', 0)
            reverse = ind_config.get('reverse', field in self.REVERSE_INDICATORS)

            # 获取指标值（fina_indicator → 别名 → 报表计算）
            value = self._resolve_indicator_value(financial_data, field)

            if value is None or pd.isna(value):
                indicator_results.append({
                    'name': ind_name,
                    'field': field,
                    'value': None,
                    'score': None,
                    'score_ratio': None,
                    'quantile_level': '数据缺失',
                    'weight': ind_weight,
                    'status': '数据缺失',
                })
                continue

            quantile_level, score_ratio = self._evaluate_indicator(
                field, value, thresholds, reverse
            )
            score = score_ratio * 100

            indicator_results.append({
                'name': ind_name,
                'field': field,
                'value': round(value, 4),
                'score': round(score, 2),
                'score_ratio': score_ratio,
                'quantile_level': quantile_level,
                'weight': ind_weight,
                'status': '正常',
            })

        # 加权平均
        if total_indicator_weight > 0:
            weighted_sum = sum(
                r['score'] * r['weight'] for r in indicator_results if r['score'] is not None
            )
            valid_weight = sum(
                r['weight'] for r in indicator_results if r['score'] is not None
            )
            dimension_score = weighted_sum / valid_weight if valid_weight > 0 else 0
        else:
            valid_scores = [r['score'] for r in indicator_results if r['score'] is not None]
            dimension_score = np.mean(valid_scores) if valid_scores else 0

        return dimension_score, indicator_results

    def _calculate_solvency_dimension(
        self,
        dim_config: Dict[str, Any],
        financial_data: pd.DataFrame,
        thresholds: Dict[str, Dict[str, float]],
        dim_weight: float,
    ) -> tuple:
        """计算偿债能力维度（含短期/长期子维度）"""
        sub_dims = dim_config.get('sub_dimensions', {})
        all_indicators = []
        sub_dim_results = {}
        sub_dim_scores = {}

        for sub_key, sub_config in sub_dims.items():
            sub_name = sub_config.get('name', sub_key)
            sub_weight = sub_config.get('weight', 0)
            indicators = sub_config.get('indicators', [])

            sub_indicators = []
            total_ind_weight = sum(ind.get('weight', 0) for ind in indicators)

            for ind_config in indicators:
                field = ind_config.get('field')
                ind_name = ind_config.get('name', self.INDICATOR_NAMES.get(field, field))
                ind_weight = ind_config.get('weight', 0)
                reverse = ind_config.get('reverse', field in self.REVERSE_INDICATORS)

                value = self._resolve_indicator_value(financial_data, field)

                if value is None or pd.isna(value):
                    result_item = {
                        'name': ind_name,
                        'field': field,
                        'value': None,
                        'score': None,
                        'score_ratio': None,
                        'quantile_level': '数据缺失',
                        'weight': ind_weight,
                        'sub_dimension': sub_name,
                        'status': '数据缺失',
                    }
                    sub_indicators.append(result_item)
                    all_indicators.append(result_item)
                    continue

                quantile_level, score_ratio = self._evaluate_indicator(
                    field, value, thresholds, reverse
                )
                score = score_ratio * 100

                result_item = {
                    'name': ind_name,
                    'field': field,
                    'value': round(value, 4),
                    'score': round(score, 2),
                    'score_ratio': score_ratio,
                    'quantile_level': quantile_level,
                    'weight': ind_weight,
                    'sub_dimension': sub_name,
                    'status': '正常',
                }
                sub_indicators.append(result_item)
                all_indicators.append(result_item)

            # 子维度加权平均
            if total_ind_weight > 0:
                weighted_sum = sum(
                    r['score'] * r['weight'] for r in sub_indicators if r['score'] is not None
                )
                valid_weight = sum(
                    r['weight'] for r in sub_indicators if r['score'] is not None
                )
                sub_score = weighted_sum / valid_weight if valid_weight > 0 else 0
            else:
                valid_scores = [r['score'] for r in sub_indicators if r['score'] is not None]
                sub_score = np.mean(valid_scores) if valid_scores else 0

            sub_dim_scores[sub_key] = sub_score
            sub_dim_results[sub_key] = {
                'name': sub_name,
                'weight': sub_weight,
                'score': round(sub_score, 2),
                'indicators': sub_indicators,
            }

        # 偿债能力总得分 = 短期偿债得分 × 短期权重 + 长期偿债得分 × 长期权重
        dimension_score = sum(
            sub_dim_scores.get(k, 0) * sub_dims[k].get('weight', 0)
            for k in sub_dims
        )

        return dimension_score, all_indicators, sub_dim_results

    def _evaluate_indicator(
        self,
        field: str,
        value: float,
        thresholds: Dict[str, Dict[str, float]],
        reverse: bool,
    ) -> tuple:
        """评估单个指标的等级和得分比例

        Returns:
            (等级名称, 得分比例)
        """
        # 特殊处理：反向指标负值直接极差
        if reverse and value < 0:
            return '极差', self.SCORE_RATIOS['极差']

        threshold = thresholds.get(field)
        if threshold is None:
            # 无阈值时使用线性回退
            score = self._linear_score(value, field, reverse)
            if score >= 90:
                return '优秀', self.SCORE_RATIOS['优秀']
            elif score >= 80:
                return '良好', self.SCORE_RATIOS['良好']
            elif score >= 70:
                return '中等', self.SCORE_RATIOS['中等']
            elif score >= 60:
                return '较差', self.SCORE_RATIOS['较差']
            else:
                return '极差', self.SCORE_RATIOS['极差']

        if reverse:
            # 反向指标：值越小越好
            if value <= threshold['p10']:
                level = '优秀'
            elif value <= threshold['p30']:
                level = '良好'
            elif value <= threshold['p50']:
                level = '中等'
            elif value <= threshold['p70']:
                level = '较差'
            else:
                level = '极差'
        else:
            # 正向指标：值越大越好
            if value >= threshold['p90']:
                level = '优秀'
            elif value >= threshold['p70']:
                level = '良好'
            elif value >= threshold['p50']:
                level = '中等'
            elif value >= threshold['p30']:
                level = '较差'
            else:
                level = '极差'

        return level, self.SCORE_RATIOS[level]

    def _build_industry_thresholds(self) -> Dict[str, Dict[str, float]]:
        """构建行业分位数阈值

        从行业数据计算 p10/p30/p50/p70/p90，无行业数据时使用默认阈值
        """
        # 所有需要阈值的指标
        all_fields = set()
        for dim_key, dim_config in self.scoring_config.items():
            if 'sub_dimensions' in dim_config:
                for sub_config in dim_config['sub_dimensions'].values():
                    for ind in sub_config.get('indicators', []):
                        all_fields.add(ind.get('field'))
            else:
                for ind in dim_config.get('indicators', []):
                    all_fields.add(ind.get('field'))

        thresholds = {}
        for field in all_fields:
            thresholds[field] = self._get_threshold_for_field(field)

        return thresholds

    def _get_threshold_for_field(self, field: str) -> Optional[Dict[str, float]]:
        """获取单个指标的阈值"""
        # 优先从真实行业数据计算
        if self._industry_data is not None and not self._industry_data.empty:
            col = field
            if col not in self._industry_data.columns:
                # 尝试别名
                col = self.FIELD_ALIASES.get(field, field)

            if col in self._industry_data.columns:
                values = self._industry_data[col].dropna().values
                if len(values) >= 5:
                    return {
                        'p10': float(np.percentile(values, 10)),
                        'p30': float(np.percentile(values, 30)),
                        'p50': float(np.percentile(values, 50)),
                        'p70': float(np.percentile(values, 70)),
                        'p90': float(np.percentile(values, 90)),
                    }

        # 默认阈值
        return self._get_default_threshold(field)

    def _get_default_threshold(self, field: str) -> Optional[Dict[str, float]]:
        """各指标的默认行业阈值（p10/p30/p50/p70/p90）"""
        defaults = {
            # 运营能力
            'inv_turn': {'p10': 2, 'p30': 4, 'p50': 6, 'p70': 8, 'p90': 12},
            'ar_turn': {'p10': 3, 'p30': 6, 'p50': 9, 'p70': 12, 'p90': 15},
            'ca_turn': {'p10': 0.5, 'p30': 0.8, 'p50': 1.2, 'p70': 1.8, 'p90': 2.5},
            'assets_turn': {'p10': 0.3, 'p30': 0.5, 'p50': 0.7, 'p70': 1.0, 'p90': 1.5},
            # 盈利能力
            'netprofit_margin': {'p10': 2, 'p30': 5, 'p50': 8, 'p70': 12, 'p90': 20},
            'grossprofit_margin': {'p10': 15, 'p30': 20, 'p50': 28, 'p70': 38, 'p90': 50},
            'roe': {'p10': 3, 'p30': 6, 'p50': 10, 'p70': 15, 'p90': 22},
            'roe_dt': {'p10': 2, 'p30': 5, 'p50': 8, 'p70': 13, 'p90': 20},
            'roa': {'p10': 1.5, 'p30': 3, 'p50': 5, 'p70': 8, 'p90': 12},
            'npta': {'p10': 1, 'p30': 3, 'p50': 5, 'p70': 7, 'p90': 10},
            # 短期偿债
            'current_ratio': {'p10': 0.8, 'p30': 1.0, 'p50': 1.5, 'p70': 2.0, 'p90': 3.0},
            'quick_ratio': {'p10': 0.5, 'p30': 0.7, 'p50': 1.0, 'p70': 1.5, 'p90': 2.5},
            'cash_to_liqdebt': {'p10': 0.1, 'p30': 0.2, 'p50': 0.4, 'p70': 0.6, 'p90': 0.8},
            'cash_to_liqdebt_withinterest': {'p10': 0.1, 'p30': 0.15, 'p50': 0.3, 'p70': 0.5, 'p90': 0.7},
            # 长期偿债（反向指标，阈值方向相反：p10是最优值，p90是最差值）
            'debt_to_assets': {'p10': 30, 'p30': 40, 'p50': 50, 'p70': 60, 'p90': 75},
            'int_to_talcap': {'p10': 15, 'p30': 25, 'p50': 35, 'p70': 45, 'p90': 60},
            'debt_to_eqt': {'p10': 0.8, 'p30': 1.2, 'p50': 1.8, 'p70': 2.5, 'p90': 4.0},
            'ebit_to_interest': {'p10': 1.5, 'p30': 2.5, 'p50': 4.0, 'p70': 6.0, 'p90': 10.0},
            # 成长能力
            'netprofit_yoy': {'p10': -15, 'p30': 0, 'p50': 10, 'p70': 25, 'p90': 45},
            'dt_netprofit_yoy': {'p10': -20, 'p30': -5, 'p50': 8, 'p70': 20, 'p90': 40},
            'roe_yoy': {'p10': -10, 'p30': 0, 'p50': 5, 'p70': 15, 'p90': 30},
            'tr_yoy': {'p10': -10, 'p30': 0, 'p50': 10, 'p70': 20, 'p90': 35},
            'or_yoy': {'p10': -10, 'p30': 0, 'p50': 10, 'p70': 20, 'p90': 35},
            'equity_yoy': {'p10': -5, 'p30': 2, 'p50': 8, 'p70': 15, 'p90': 25},
            'op_yoy': {'p10': -15, 'p30': 0, 'p50': 10, 'p70': 25, 'p90': 40},
            'ebt_yoy': {'p10': -15, 'p30': 0, 'p50': 10, 'p70': 20, 'p90': 40},
            'rd_exp_ratio': {'p10': 1, 'p30': 3, 'p50': 5, 'p70': 8, 'p90': 12},
            # 兼容旧字段
            'net_margin': {'p10': 2, 'p30': 5, 'p50': 8, 'p70': 12, 'p90': 20},
            'gross_margin': {'p10': 15, 'p30': 20, 'p50': 28, 'p70': 38, 'p90': 50},
            'arturn_days': {'p10': 30, 'p30': 45, 'p50': 60, 'p70': 90, 'p90': 150},
            'current_asset_turn': {'p10': 0.5, 'p30': 0.8, 'p50': 1.2, 'p70': 1.8, 'p90': 2.5},
            'asset_turn': {'p10': 0.3, 'p30': 0.5, 'p50': 0.7, 'p70': 1.0, 'p90': 1.5},
            'net_profit_yoy': {'p10': -15, 'p30': 0, 'p50': 10, 'p70': 25, 'p90': 45},
            'revenue_yoy': {'p10': -10, 'p30': 0, 'p50': 10, 'p70': 20, 'p90': 35},
        }
        return defaults.get(field)

    def _resolve_indicator_value(self, financial_data: pd.DataFrame, field: str) -> Optional[float]:
        """按优先级获取指标值：fina_indicator → 别名 → 从报表计算"""
        # 1. 直接从fina_indicator获取
        value = self._get_latest_value(financial_data, field)
        if value is not None:
            return value

        # 2. 尝试别名
        alias = self.FIELD_ALIASES.get(field)
        if alias:
            value = self._get_latest_value(financial_data, alias)
            if value is not None:
                return value

        # 反向查找别名
        for orig, a in self.FIELD_ALIASES.items():
            if a == field:
                value = self._get_latest_value(financial_data, orig)
                if value is not None:
                    return value

        # 3. 从原始财务报表计算
        value = self._calculate_from_statements(field)
        return value

    def _get_latest_value(self, df: pd.DataFrame, field: str) -> Optional[float]:
        """获取最新一期有效数据"""
        if field not in df.columns:
            return None

        df_sorted = df.sort_values('end_date', ascending=False)
        for _, row in df_sorted.iterrows():
            value = row.get(field)
            if value is not None and not pd.isna(value):
                return float(value)
        return None

    def _calculate_from_statements(self, field: str) -> Optional[float]:
        """从原始财务报表计算缺失指标

        Tushare资产负债表/利润表单位为元，fina_indicator中比率类为百分数或倍数。
        计算结果需与fina_indicator保持一致格式。
        """
        if self._financial_statements is None:
            return None

        balance = self._financial_statements.get('balance_sheet', pd.DataFrame())
        income = self._financial_statements.get('income_statement', pd.DataFrame())

        if balance.empty and income.empty:
            return None

        # 获取最新一期资产负债表
        bs = self._get_latest_period(balance)
        ic = self._get_latest_period(income)

        if bs is None and ic is None:
            return None

        try:
            return self._do_calculate(field, bs, ic, balance, income)
        except (ZeroDivisionError, TypeError, ValueError):
            return None

    def _get_latest_period(self, df: pd.DataFrame) -> Optional[pd.Series]:
        """获取最新一期的数据行"""
        if df.empty:
            return None
        df_sorted = df.sort_values('end_date', ascending=False)
        for _, row in df_sorted.iterrows():
            return row
        return None

    def _safe_float(self, value) -> Optional[float]:
        """安全转换为float"""
        if value is None or pd.isna(value):
            return None
        return float(value)

    def _do_calculate(self, field: str, bs: Optional[pd.Series], ic: Optional[pd.Series],
                      balance_df: pd.DataFrame, income_df: pd.DataFrame) -> Optional[float]:
        """执行具体的指标计算"""

        # === 运营能力 ===
        if field == 'inv_turn':
            # 存货周转率 = 营业成本 / 平均存货
            if ic is None or bs is None:
                return None
            oper_cost = self._safe_float(ic.get('oper_cost'))
            if oper_cost is None:
                return None
            invty = self._safe_float(bs.get('invty'))
            if invty is None or invty == 0:
                return None
            return oper_cost / invty

        if field == 'ar_turn':
            # 应收账款周转率 = 营业收入 / 平均应收账款
            if ic is None or bs is None:
                return None
            revenue = self._safe_float(ic.get('total_revenue')) or self._safe_float(ic.get('revenue'))
            ar = self._safe_float(bs.get('account_rece'))
            if revenue is None or ar is None or ar == 0:
                return None
            return revenue / ar

        if field in ('ca_turn', 'current_asset_turn'):
            # 流动资产周转率 = 营业收入 / 平均流动资产
            if ic is None or bs is None:
                return None
            revenue = self._safe_float(ic.get('total_revenue')) or self._safe_float(ic.get('revenue'))
            cur_assets = self._safe_float(bs.get('total_cur_assets'))
            if revenue is None or cur_assets is None or cur_assets == 0:
                return None
            return revenue / cur_assets

        if field in ('assets_turn', 'asset_turn'):
            # 总资产周转率 = 营业收入 / 平均总资产
            if ic is None or bs is None:
                return None
            revenue = self._safe_float(ic.get('total_revenue')) or self._safe_float(ic.get('revenue'))
            total_assets = self._safe_float(bs.get('total_assets'))
            if revenue is None or total_assets is None or total_assets == 0:
                return None
            return revenue / total_assets

        # === 盈利能力 ===
        if field in ('netprofit_margin', 'net_margin'):
            if ic is None:
                return None
            net_income = self._safe_float(ic.get('n_income_attr_p')) or self._safe_float(ic.get('n_income'))
            revenue = self._safe_float(ic.get('total_revenue')) or self._safe_float(ic.get('revenue'))
            if net_income is None or revenue is None or revenue == 0:
                return None
            return net_income / revenue * 100

        if field in ('grossprofit_margin', 'gross_margin'):
            if ic is None:
                return None
            revenue = self._safe_float(ic.get('total_revenue')) or self._safe_float(ic.get('revenue'))
            oper_cost = self._safe_float(ic.get('oper_cost'))
            if revenue is None or oper_cost is None or revenue == 0:
                return None
            return (revenue - oper_cost) / revenue * 100

        if field == 'roe':
            if ic is None or bs is None:
                return None
            net_income = self._safe_float(ic.get('n_income_attr_p')) or self._safe_float(ic.get('n_income'))
            equity = self._safe_float(bs.get('total_hldr_eqy_exc_min_int'))
            if net_income is None or equity is None or equity == 0:
                return None
            return net_income / equity * 100

        if field == 'roa':
            if ic is None or bs is None:
                return None
            net_income = self._safe_float(ic.get('n_income'))
            total_assets = self._safe_float(bs.get('total_assets'))
            if net_income is None or total_assets is None or total_assets == 0:
                return None
            return net_income / total_assets * 100

        if field == 'npta':
            if ic is None or bs is None:
                return None
            net_income = self._safe_float(ic.get('n_income'))
            total_assets = self._safe_float(bs.get('total_assets'))
            if net_income is None or total_assets is None or total_assets == 0:
                return None
            return net_income / total_assets * 100

        # === 短期偿债 ===
        if field == 'current_ratio':
            if bs is None:
                return None
            cur_assets = self._safe_float(bs.get('total_cur_assets'))
            cur_liab = self._safe_float(bs.get('total_cur_liab'))
            if cur_assets is None or cur_liab is None or cur_liab == 0:
                return None
            return cur_assets / cur_liab

        if field == 'quick_ratio':
            if bs is None:
                return None
            cur_assets = self._safe_float(bs.get('total_cur_assets'))
            invty = self._safe_float(bs.get('invty')) or 0
            cur_liab = self._safe_float(bs.get('total_cur_liab'))
            if cur_assets is None or cur_liab is None or cur_liab == 0:
                return None
            return (cur_assets - invty) / cur_liab

        if field == 'cash_to_liqdebt':
            # 现金与流动负债比率 = 货币资金 / 流动负债
            if bs is None:
                return None
            money_cap = self._safe_float(bs.get('money_cap'))
            cur_liab = self._safe_float(bs.get('total_cur_liab'))
            if money_cap is None or cur_liab is None or cur_liab == 0:
                return None
            return money_cap / cur_liab

        if field == 'cash_to_liqdebt_withinterest':
            # 现金与带息流动负债比率 = 货币资金 / (短期借款 + 一年内到期非流动负债)
            if bs is None:
                return None
            money_cap = self._safe_float(bs.get('money_cap'))
            st_loan = self._safe_float(bs.get('st_loan')) or 0
            non_cur_1y = self._safe_float(bs.get('non_cur_liab_within_1y')) or 0
            denom = st_loan + non_cur_1y
            if money_cap is None or denom == 0:
                return None
            return money_cap / denom

        # === 长期偿债 ===
        if field == 'debt_to_assets':
            if bs is None:
                return None
            total_liab = self._safe_float(bs.get('total_liab'))
            total_assets = self._safe_float(bs.get('total_assets'))
            if total_liab is None or total_assets is None or total_assets == 0:
                return None
            return total_liab / total_assets * 100

        if field == 'int_to_talcap':
            # 带息债务/全部投入资本 = 有息负债 / (有息负债 + 股东权益)
            if bs is None:
                return None
            st_loan = self._safe_float(bs.get('st_loan')) or 0
            lt_borr = self._safe_float(bs.get('lt_borr')) or 0
            bonds = self._safe_float(bs.get('bonds_payable')) or 0
            non_cur_1y = self._safe_float(bs.get('non_cur_liab_within_1y')) or 0
            interest_debt = st_loan + lt_borr + bonds + non_cur_1y
            equity = self._safe_float(bs.get('total_hldr_eqy_exc_min_int'))
            if equity is None:
                return None
            denom = interest_debt + equity
            if denom == 0:
                return None
            return interest_debt / denom * 100

        if field == 'debt_to_eqt':
            # 产权比率 = 负债合计 / 股东权益
            if bs is None:
                return None
            total_liab = self._safe_float(bs.get('total_liab'))
            equity = self._safe_float(bs.get('total_hldr_eqy_exc_min_int'))
            if total_liab is None or equity is None or equity == 0:
                return None
            return total_liab / equity

        if field == 'ebit_to_interest':
            # 利息保障倍数 = EBIT / 利息费用
            if ic is None:
                return None
            ebit = self._safe_float(ic.get('ebit'))
            interest = self._safe_float(ic.get('interest_exp'))
            if ebit is None:
                return None
            if interest is None or interest == 0:
                # 无利息费用时，说明偿债压力小，给高倍数
                return None
            return ebit / abs(interest)

        # === 成长能力（需要同比计算，从两年数据对比） ===
        if field in ('netprofit_yoy', 'net_profit_yoy'):
            return self._calc_yoy(income_df, 'n_income_attr_p', 'n_income')

        if field == 'dt_netprofit_yoy':
            return self._calc_yoy(income_df, 'n_income_attr_p')

        if field in ('or_yoy', 'revenue_yoy'):
            return self._calc_yoy(income_df, 'total_revenue', 'revenue')

        if field == 'tr_yoy':
            return self._calc_yoy(income_df, 'total_revenue')

        if field == 'op_yoy':
            return self._calc_yoy(income_df, 'operate_profit')

        if field == 'ebt_yoy':
            return self._calc_yoy(income_df, 'total_profit')

        if field == 'equity_yoy':
            return self._calc_yoy(balance_df, 'total_hldr_eqy_exc_min_int')

        if field == 'roe_yoy':
            # ROE同比 = 今年ROE - 去年ROE
            roe_this = self._calc_roe_for_year(income_df, balance_df, 0)
            roe_prev = self._calc_roe_for_year(income_df, balance_df, -1)
            if roe_this is not None and roe_prev is not None and roe_prev != 0:
                return (roe_this - roe_prev) / abs(roe_prev) * 100
            return None

        if field == 'rd_exp_ratio':
            # 研发费用率 = 研发费用 / 营业收入 × 100
            if ic is None:
                return None
            rd_exp = self._safe_float(ic.get('rd_exp'))
            revenue = self._safe_float(ic.get('total_revenue')) or self._safe_float(ic.get('revenue'))
            if rd_exp is None or revenue is None or revenue == 0:
                return None
            return rd_exp / revenue * 100

        return None

    def _calc_yoy(self, df: pd.DataFrame, *fields) -> Optional[float]:
        """从连续两年数据计算同比增长率"""
        if df.empty or 'end_date' not in df.columns:
            return None
        df_sorted = df.sort_values('end_date', ascending=False)
        annual = df_sorted[df_sorted['end_date'].astype(str).str.contains('1231', na=False)]

        if len(annual) < 2:
            return None

        # 获取当年和上年值
        for field in fields:
            curr = self._safe_float(annual.iloc[0].get(field))
            prev = self._safe_float(annual.iloc[1].get(field))
            if curr is not None and prev is not None and prev != 0:
                return (curr - prev) / abs(prev) * 100
        return None

    def _calc_roe_for_year(self, income_df: pd.DataFrame, balance_df: pd.DataFrame,
                           offset: int) -> Optional[float]:
        """计算指定年份的ROE（offset=0当年，offset=-1上年）"""
        if income_df.empty or balance_df.empty:
            return None

        inc_sorted = income_df.sort_values('end_date', ascending=False)
        bs_sorted = balance_df.sort_values('end_date', ascending=False)

        inc_annual = inc_sorted[inc_sorted['end_date'].astype(str).str.contains('1231', na=False)]
        bs_annual = bs_sorted[bs_sorted['end_date'].astype(str).str.contains('1231', na=False)]

        idx_inc = -offset if offset < 0 else 0
        idx_bs = -offset if offset < 0 else 0

        if len(inc_annual) <= idx_inc or len(bs_annual) <= idx_bs:
            return None

        net_income = self._safe_float(inc_annual.iloc[idx_inc].get('n_income_attr_p'))
        equity = self._safe_float(bs_annual.iloc[idx_bs].get('total_hldr_eqy_exc_min_int'))

        if net_income is None or equity is None or equity == 0:
            return None
        return net_income / equity * 100

    def _linear_score(self, value: float, field: str, reverse: bool = False) -> float:
        """线性回退评分"""
        ranges = {
            'inv_turn': (0, 15),
            'ar_turn': (0, 15),
            'ca_turn': (0, 3),
            'assets_turn': (0, 3),
            'netprofit_margin': (-5, 25),
            'grossprofit_margin': (0, 50),
            'roe': (-5, 30),
            'roe_dt': (-5, 30),
            'roa': (-5, 20),
            'npta': (-5, 15),
            'current_ratio': (0.3, 4.0),
            'quick_ratio': (0.1, 3.0),
            'cash_to_liqdebt': (0, 1.0),
            'cash_to_liqdebt_withinterest': (0, 1.0),
            'debt_to_assets': (10, 80),
            'int_to_talcap': (0, 70),
            'debt_to_eqt': (0, 5.0),
            'ebit_to_interest': (-2, 20),
            'netprofit_yoy': (-30, 50),
            'dt_netprofit_yoy': (-30, 50),
            'roe_yoy': (-20, 40),
            'tr_yoy': (-30, 50),
            'or_yoy': (-30, 50),
            'equity_yoy': (-20, 40),
            'op_yoy': (-30, 50),
            'ebt_yoy': (-30, 50),
            'rd_exp_ratio': (0, 15),
        }

        if field in ranges:
            low, high = ranges[field]
            if high == low:
                return 50
            score = (value - low) / (high - low) * 100
            if reverse:
                score = 100 - score
            return max(0, min(100, score))

        if reverse:
            return max(0, 100 - value)
        return min(100, value)

    def _get_grade(self, score: float) -> str:
        for grade, (low, high) in self.GRADE_LEVELS.items():
            if low <= score <= high:
                return grade
        return '极差'

    def _empty_result(self) -> Dict[str, Any]:
        return {
            'total_score': 0,
            'grade': '暂无数据',
            'dimensions': {},
            'indicators': {},
            'sub_dimensions': {},
        }

    def get_dimension_summary(self, result: Dict[str, Any]) -> pd.DataFrame:
        """获取维度评分汇总表"""
        rows = []
        for dim, data in result.get('dimensions', {}).items():
            rows.append({
                '维度': self.DIMENSION_NAMES.get(dim, dim),
                '权重': f"{data.get('weight', 0) * 100:.0f}%",
                '得分': data.get('score', 0),
                '加权得分': round(data.get('weighted_score', 0), 2),
            })

        df = pd.DataFrame(rows)
        if not df.empty:
            df.loc['合计'] = ['总分', '100%', result.get('total_score', 0), result.get('total_score', 0)]
        return df

    def get_indicator_detail_table(self, result: Dict[str, Any]) -> pd.DataFrame:
        """获取指标详情表"""
        rows = []
        for dim, indicators in result.get('indicators', {}).items():
            dim_name = self.DIMENSION_NAMES.get(dim, dim)
            for ind in indicators:
                rows.append({
                    '维度': dim_name,
                    '子维度': ind.get('sub_dimension', ''),
                    '指标名称': ind.get('name', ''),
                    '指标值': ind.get('value'),
                    '等级': ind.get('quantile_level', ''),
                    '得分比例': ind.get('score_ratio'),
                    '指标得分': ind.get('score'),
                    '状态': ind.get('status', ''),
                })
        return pd.DataFrame(rows)

    def get_scoring_summary_text(self, result: Dict[str, Any]) -> str:
        """生成评分总结文本"""
        total = result.get('total_score', 0)
        grade = result.get('grade', '暂无数据')

        summary = f"综合财务评分: {total:.2f}分\n"
        summary += f"评级: {grade}\n\n"

        summary += "各维度评分:\n"
        for dim, data in result.get('dimensions', {}).items():
            dim_name = self.DIMENSION_NAMES.get(dim, dim)
            score = data.get('score', 0)
            weight = data.get('weight', 0)
            summary += f"  - {dim_name}(权重{weight * 100:.0f}%): {score:.2f}分\n"

        return summary
