"""
情景分析模块
包含基准情景、乐观情景、悲观情景等多情景分析
"""
import copy
from typing import Dict, List, Optional, Any
from models import Company, ValuationResult, ScenarioConfig, SCENARIOS
from absolute_valuation import AbsoluteValuation


class ScenarioAnalyzer:
    """情景分析器"""

    def __init__(self, company: Company, base_valuation: Optional[ValuationResult] = None):
        """
        初始化情景分析器

        Args:
            company: 目标公司
            base_valuation: 基准估值结果
        """
        self.company = company
        self.base_valuation = base_valuation

    def base_case(
        self,
        valuation_method: str = "DCF",
        method_params: Optional[Dict[str, Any]] = None
    ) -> ValuationResult:
        """
        基准情景分析

        基于管理层预测或当前共识进行估值

        Args:
            valuation_method: 估值方法（DCF、PE等）
            method_params: 方法参数

        Returns:
            基准情景估值结果
        """
        if valuation_method == "DCF":
            params = method_params or {}
            return AbsoluteValuation.dcf_valuation(self.company, **params)
        else:
            raise ValueError(f"暂不支持{valuation_method}方法")

    def bull_case(
        self,
        revenue_boost: float = 0.2,
        margin_boost: float = 0.05,
        wadj_reduction: float = 0.01,
        terminal_growth_boost: float = 0.005,
        valuation_method: str = "DCF",
        method_params: Optional[Dict[str, Any]] = None
    ) -> ValuationResult:
        """
        乐观情景分析

        假设收入增长更快、利润率提升、风险降低

        Args:
            revenue_boost: 收入增长率提升幅度
            margin_boost: 利润率提升幅度
            wacc_reduction: WACC降低幅度
            terminal_growth_boost: 永续增长率提升幅度
            valuation_method: 估值方法
            method_params: 方法参数

        Returns:
            乐观情景估值结果
        """
        # 创建调整后的公司参数
        custom_assumptions = {
            'growth_rate': self.company.growth_rate + revenue_boost,
            'operating_margin': (self.company.operating_margin or 0.2) + margin_boost,
        }

        params = method_params or {}
        params['custom_assumptions'] = custom_assumptions
        params['wacc'] = params.get('wacc', 0) - wacc_reduction if 'wacc' in params else None
        params['terminal_growth_rate'] = params.get('terminal_growth_rate',
                                                      self.company.terminal_growth_rate) + terminal_growth_boost

        if valuation_method == "DCF":
            return AbsoluteValuation.dcf_valuation(self.company, **params)
        else:
            raise ValueError(f"暂不支持{valuation_method}方法")

    def bear_case(
        self,
        revenue_reduction: float = 0.2,
        margin_reduction: float = 0.05,
        wacc_increase: float = 0.02,
        terminal_growth_reduction: float = 0.005,
        valuation_method: str = "DCF",
        method_params: Optional[Dict[str, Any]] = None
    ) -> ValuationResult:
        """
        悲观情景分析

        假设增长放缓、利润率下降、风险上升

        Args:
            revenue_reduction: 收入增长率降低幅度
            margin_reduction: 利润率降低幅度
            wacc_increase: WACC增加幅度
            terminal_growth_reduction: 永续增长率降低幅度
            valuation_method: 估值方法
            method_params: 方法参数

        Returns:
            悲观情景估值结果
        """
        # 确保不会出现负数
        new_growth = max(0, self.company.growth_rate - revenue_reduction)
        new_margin = max(0, (self.company.operating_margin or 0.2) - margin_reduction)

        custom_assumptions = {
            'growth_rate': new_growth,
            'operating_margin': new_margin,
        }

        params = method_params or {}
        params['custom_assumptions'] = custom_assumptions
        params['wacc'] = params.get('wacc', 0) + wacc_increase if 'wacc' in params else None
        params['terminal_growth_rate'] = max(0, params.get('terminal_growth_rate',
                                                             self.company.terminal_growth_rate) - terminal_growth_reduction)

        if valuation_method == "DCF":
            return AbsoluteValuation.dcf_valuation(self.company, **params)
        else:
            raise ValueError(f"暂不支持{valuation_method}方法")

    def custom_scenario(
        self,
        scenario: ScenarioConfig,
        valuation_method: str = "DCF",
        method_params: Optional[Dict[str, Any]] = None
    ) -> ValuationResult:
        """
        自定义情景分析

        Args:
            scenario: 情景配置
            valuation_method: 估值方法
            method_params: 方法参数

        Returns:
            自定义情景估值结果
        """
        custom_assumptions = {
            'growth_rate': self.company.growth_rate + scenario.revenue_growth_adj,
            'operating_margin': (self.company.operating_margin or 0.2) + scenario.margin_adj,
        }

        params = method_params or {}
        params['custom_assumptions'] = custom_assumptions
        params['wacc'] = params.get('wacc', 0) + scenario.wacc_adj if 'wacc' in params else None
        params['terminal_growth_rate'] = params.get('terminal_growth_rate',
                                                      self.company.terminal_growth_rate) + scenario.terminal_growth_adj

        if valuation_method == "DCF":
            return AbsoluteValuation.dcf_valuation(self.company, **params)
        else:
            raise ValueError(f"暂不支持{valuation_method}方法")

    def compare_scenarios(
        self,
        scenarios: Optional[List[ScenarioConfig]] = None,
        valuation_method: str = "DCF",
        method_params: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        多情景对比分析

        Args:
            scenarios: 情景列表（如不提供则使用默认三情景）
            valuation_method: 估值方法
            method_params: 方法参数

        Returns:
            情景对比结果
        """
        if scenarios is None:
            scenarios = [SCENARIOS['base'], SCENARIOS['bull'], SCENARIOS['bear']]

        results = {}

        for scenario in scenarios:
            try:
                result = self.custom_scenario(scenario, valuation_method, method_params)
                results[scenario.name] = {
                    'scenario': scenario,
                    'valuation': result,
                    'value': result.value,
                }
            except Exception as e:
                print(f"情景'{scenario.name}'计算失败: {e}")

        # 计算统计信息
        values = [r['value'] for r in results.values()]
        if values:
            import numpy as np
            results['statistics'] = {
                'mean': np.mean(values),
                'median': np.median(values),
                'std': np.std(values),
                'min': np.min(values),
                'max': np.max(values),
                'range': np.max(values) - np.min(values),
                'count': len(values),
            }

        return results

    def scenario_probability_analysis(
        self,
        scenarios_with_probability: List[tuple],
        valuation_method: str = "DCF",
        method_params: Optional[Dict[str, Any]] = None
    ) -> ValuationResult:
        """
        情景概率分析

        基于不同情景的发生概率计算期望价值

        Args:
            scenarios_with_probability: (情景, 概率) 元组列表
            valuation_method: 估值方法
            method_params: 方法参数

        Returns:
            期望估值结果
        """
        total_probability = 0
        expected_value = 0

        scenario_values = []

        for scenario, probability in scenarios_with_probability:
            result = self.custom_scenario(scenario, valuation_method, method_params)
            expected_value += result.value * probability
            total_probability += probability
            scenario_values.append({
                'scenario': scenario.name,
                'probability': probability,
                'value': result.value,
                'contribution': result.value * probability,
            })

        # 标准化
        if total_probability > 0:
            expected_value /= total_probability

        return ValuationResult(
            method="情景概率分析",
            value=expected_value,
            details={
                'scenario_values': scenario_values,
                'total_probability': total_probability,
                'scenarios': scenarios_with_probability,
            },
            assumptions={}
        )


# ===== 辅助函数 =====

def create_scenario_report(
    analysis_results: Dict[str, Any],
    format: str = "text"
) -> str:
    """
    生成情景分析报告

    Args:
        analysis_results: 情景分析结果
        format: 报告格式（text/markdown/json）

    Returns:
        格式化的报告
    """
    if format == "text":
        return _create_text_report(analysis_results)
    elif format == "markdown":
        return _create_markdown_report(analysis_results)
    elif format == "json":
        import json
        return json.dumps(analysis_results, indent=2, default=str)
    else:
        raise ValueError(f"不支持的格式: {format}")


def _create_text_report(results: Dict[str, Any]) -> str:
    """创建文本格式报告"""
    output = []
    output.append("=" * 70)
    output.append(f"{'情景分析报告':^68}")
    output.append("=" * 70)

    # 各情景结果
    for name, data in results.items():
        if name == 'statistics':
            continue
        scenario = data['scenario']
        valuation = data['valuation']

        output.append(f"\n【{name}】")
        output.append(f"  参数调整:")
        output.append(f"    收入增长: {scenario.revenue_growth_adj:+.1%}")
        output.append(f"    利润率: {scenario.margin_adj:+.1%}")
        output.append(f"    WACC: {scenario.wacc_adj:+.1%}")
        output.append(f"    终值增长: {scenario.terminal_growth_adj:+.1%}")
        output.append(f"  估值: {valuation.value/10000:.2f}亿元")
        if valuation.value_low and valuation.value_high:
            output.append(f"  区间: {valuation.value_low/10000:.2f} - {valuation.value_high/10000:.2f}亿元")

    # 统计信息
    if 'statistics' in results:
        stats = results['statistics']
        output.append(f"\n【统计摘要】")
        output.append(f"  平均值: {stats['mean']/10000:.2f}亿元")
        output.append(f"  中位数: {stats['median']/10000:.2f}亿元")
        output.append(f"  标准差: {stats['std']/10000:.2f}亿元")
        output.append(f"  范围: {stats['min']/10000:.2f} - {stats['max']/10000:.2f}亿元")
        output.append(f"  波动范围: {stats['range']/10000:.2f}亿元")

    output.append("\n" + "=" * 70)

    return "\n".join(output)


def _create_markdown_report(results: Dict[str, Any]) -> str:
    """创建Markdown格式报告"""
    output = []
    output.append("# 情景分析报告\n")

    # 表格
    output.append("## 情景对比\n")
    output.append("| 情景 | 估值(亿元) | vs基准 |")
    output.append("|------|-----------|--------|")

    base_value = None
    for name, data in results.items():
        if name == 'statistics':
            continue
        value = data['value']
        if base_value is None:
            base_value = value

        diff_pct = ((value - base_value) / base_value * 100) if base_value > 0 else 0
        output.append(f"| {name} | {value/10000:.2f} | {diff_pct:+.1f}% |")

    # 统计信息
    if 'statistics' in results:
        stats = results['statistics']
        output.append("\n## 统计摘要\n")
        output.append(f"- 平均值: {stats['mean']/10000:.2f}亿元")
        output.append(f"- 中位数: {stats['median']/10000:.2f}亿元")
        output.append(f"- 标准差: {stats['std']/10000:.2f}亿元")
        output.append(f"- 估值区间: {stats['min']/10000:.2f} - {stats['max']/10000:.2f}亿元")

    return "\n".join(output)


# ===== 使用示例 =====

if __name__ == "__main__":
    from models import Company, CompanyStage

    # 创建测试公司
    company = Company(
        name="测试股份有限公司",
        industry="软件服务",
        stage=CompanyStage.GROWTH,
        revenue=50000,  # 5亿
        net_income=8000,  # 8000万
        ebitda=12000,
        net_assets=20000,
        total_debt=5000,
        cash_and_equivalents=2000,
        growth_rate=0.25,
        operating_margin=0.25,
        beta=1.2,
        terminal_growth_rate=0.025,
    )

    # 创建情景分析器
    analyzer = ScenarioAnalyzer(company)

    # 三情景分析
    print("执行三情景分析...")
    results = analyzer.compare_scenarios()

    # 生成报告
    report = create_scenario_report(results, format="text")
    print(report)

    # 概率分析
    print("\n【情景概率分析】")
    probability_scenarios = [
        (SCENARIOS['bull'], 0.2),   # 20%概率乐观
        (SCENARIOS['base'], 0.5),   # 50%概率基准
        (SCENARIOS['bear'], 0.3),   # 30%概率悲观
    ]

    expected_result = analyzer.scenario_probability_analysis(probability_scenarios)
    print(f"期望价值: {expected_result.value/10000:.2f}亿元")

    print("\n各情景贡献:")
    for item in expected_result.details['scenario_values']:
        print(f"  {item['scenario']}: {item['probability']:.0%} -> "
              f"{item['value']/10000:.2f}亿元 (贡献{item['contribution']/10000:.2f}亿元)")
