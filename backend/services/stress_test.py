"""
压力测试模块
包含收入冲击、毛利率压缩、WACC冲击等压力测试，以及蒙特卡洛模拟
"""
import numpy as np
from typing import List, Dict, Optional, Any
from core.models import Company, ValuationResult, StressTestResult, MonteCarloResult
from services.absolute_valuation import AbsoluteValuation


class StressTester:
    """压力测试器"""

    def __init__(self, company: Company, valuation_method: str = "DCF"):
        """
        初始化压力测试器

        Args:
            company: 目标公司
            valuation_method: 估值方法
        """
        self.company = company
        self.valuation_method = valuation_method

        # 计算基准估值
        self.base_valuation = self._get_base_valuation()

    def _get_base_valuation(self) -> ValuationResult:
        """获取基准估值"""
        if self.valuation_method == "DCF":
            return AbsoluteValuation.dcf_valuation(self.company)
        else:
            raise ValueError(f"暂不支持{self.valuation_method}方法")

    def revenue_shock_test(
        self,
        shocks: Optional[List[float]] = None
    ) -> List[StressTestResult]:
        """
        收入冲击测试

        Args:
            shocks: 收入变化列表（如-0.3表示下降30%）

        Returns:
            压力测试结果列表
        """
        if shocks is None:
            shocks = [-0.3, -0.2, -0.1]

        results = []
        base_value = self.base_valuation.value

        for shock in shocks:
            custom_assumptions = {
                'growth_rate': max(0, self.company.growth_rate * (1 + shock)),
            }

            stressed_valuation = AbsoluteValuation.dcf_valuation(
                self.company, custom_assumptions=custom_assumptions
            )

            change_pct = (stressed_valuation.value - base_value) / base_value if base_value > 0 else 0

            results.append(StressTestResult(
                test_name="收入冲击测试",
                scenario_description=f"收入{'下降' if shock < 0 else '上升'}{abs(shock):.0%}",
                base_value=base_value,
                stressed_value=stressed_valuation.value,
                change_pct=change_pct,
                details={
                    'shock': shock,
                    'new_growth_rate': custom_assumptions['growth_rate'],
                }
            ))

        return results

    def margin_compression_test(
        self,
        compression_levels: Optional[List[float]] = None
    ) -> List[StressTestResult]:
        """
        毛利率压缩测试

        Args:
            compression_levels: 利润率压缩幅度列表

        Returns:
            压力测试结果列表
        """
        if compression_levels is None:
            compression_levels = [0.05, 0.10, 0.15]

        results = []
        base_value = self.base_valuation.value
        base_margin = self.company.operating_margin or 0.2

        for compression in compression_levels:
            new_margin = max(0, base_margin - compression)

            custom_assumptions = {
                'operating_margin': new_margin,
            }

            stressed_valuation = AbsoluteValuation.dcf_valuation(
                self.company, custom_assumptions=custom_assumptions
            )

            change_pct = (stressed_valuation.value - base_value) / base_value if base_value > 0 else 0

            results.append(StressTestResult(
                test_name="毛利率压缩测试",
                scenario_description=f"利润率下降{compression:.0%}（从{base_margin:.1%}到{new_margin:.1%}）",
                base_value=base_value,
                stressed_value=stressed_valuation.value,
                change_pct=change_pct,
                details={
                    'compression': compression,
                    'base_margin': base_margin,
                    'new_margin': new_margin,
                }
            ))

        return results

    def wacc_shock_test(
        self,
        wacc_increases: Optional[List[float]] = None
    ) -> List[StressTestResult]:
        """
        WACC冲击测试（折现率上升）

        Args:
            wacc_increases: WACC增加幅度列表

        Returns:
            压力测试结果列表
        """
        if wacc_increases is None:
            wacc_increases = [0.01, 0.02, 0.03]

        results = []
        base_value = self.base_valuation.value

        # 获取基准WACC
        base_wacc = AbsoluteValuation.calculate_wacc(self.company)

        for increase in wacc_increases:
            new_wacc = base_wacc + increase

            stressed_valuation = AbsoluteValuation.dcf_valuation(
                self.company, wacc=new_wacc
            )

            change_pct = (stressed_valuation.value - base_value) / base_value if base_value > 0 else 0

            results.append(StressTestResult(
                test_name="WACC冲击测试",
                scenario_description=f"WACC上升{increase:.1%}（从{base_wacc:.2%}到{new_wacc:.2%}）",
                base_value=base_value,
                stressed_value=stressed_valuation.value,
                change_pct=change_pct,
                details={
                    'wacc_increase': increase,
                    'base_wacc': base_wacc,
                    'new_wacc': new_wacc,
                }
            ))

        return results

    def growth_slowdown_test(
        self,
        slowdown_factors: Optional[List[float]] = None
    ) -> List[StressTestResult]:
        """
        增长放缓测试

        Args:
            slowdown_factors: 增长放缓因子列表（如0.5表示增长率减半）

        Returns:
            压力测试结果列表
        """
        if slowdown_factors is None:
            slowdown_factors = [0.3, 0.5, 0.7]

        results = []
        base_value = self.base_valuation.value
        base_growth = self.company.growth_rate

        for factor in slowdown_factors:
            new_growth = base_growth * factor

            custom_assumptions = {
                'growth_rate': new_growth,
            }

            stressed_valuation = AbsoluteValuation.dcf_valuation(
                self.company, custom_assumptions=custom_assumptions
            )

            change_pct = (stressed_valuation.value - base_value) / base_value if base_value > 0 else 0

            results.append(StressTestResult(
                test_name="增长放缓测试",
                scenario_description=f"增长率降至{factor:.0%}（从{base_growth:.1%}到{new_growth:.1%}）",
                base_value=base_value,
                stressed_value=stressed_valuation.value,
                change_pct=change_pct,
                details={
                    'slowdown_factor': factor,
                    'base_growth': base_growth,
                    'new_growth': new_growth,
                }
            ))

        return results

    def extreme_market_crash(self) -> StressTestResult:
        """
        极端市场崩溃情景

        综合考虑收入下降、利润率压缩、WACC上升等多重负面因素

        Returns:
            压力测试结果
        """
        base_value = self.base_valuation.value

        # 极端情景参数
        revenue_decline = -0.40  # 收入下降40%
        margin_compression = 0.10  # 利润率压缩10%
        wacc_increase = 0.03  # WACC上升3%

        custom_assumptions = {
            'growth_rate': max(0, self.company.growth_rate * (1 + revenue_decline)),
            'operating_margin': max(0, (self.company.operating_margin or 0.2) - margin_compression),
        }

        base_wacc = AbsoluteValuation.calculate_wacc(self.company)

        stressed_valuation = AbsoluteValuation.dcf_valuation(
            self.company,
            custom_assumptions=custom_assumptions,
            wacc=base_wacc + wacc_increase
        )

        change_pct = (stressed_valuation.value - base_value) / base_value if base_value > 0 else 0

        return StressTestResult(
            test_name="极端市场崩溃测试",
            scenario_description=f"收入{revenue_decline:.0%}，利润率-{margin_compression:.0%}，WACC+{wacc_increase:.0%}",
            base_value=base_value,
            stressed_value=stressed_valuation.value,
            change_pct=change_pct,
            details={
                'revenue_decline': revenue_decline,
                'margin_compression': margin_compression,
                'wacc_increase': wacc_increase,
                'new_growth_rate': custom_assumptions['growth_rate'],
                'new_margin': custom_assumptions['operating_margin'],
                'new_wacc': base_wacc + wacc_increase,
            }
        )

    def monte_carlo_simulation(
        self,
        iterations: int = 1000,
        seed: Optional[int] = None
    ) -> MonteCarloResult:
        """
        蒙特卡洛模拟

        对关键参数进行随机采样，生成估值分布

        Args:
            iterations: 迭代次数
            seed: 随机种子

        Returns:
            蒙特卡洛模拟结果
        """
        if seed:
            np.random.seed(seed)

        values = []

        # 定义参数分布
        growth_rate_mean = self.company.growth_rate
        growth_rate_std = 0.05  # 增长率标准差

        margin_mean = self.company.operating_margin or 0.2
        margin_std = 0.03  # 利润率标准差

        wacc_mean = AbsoluteValuation.calculate_wacc(self.company)
        wacc_std = 0.01  # WACC标准差

        terminal_growth_mean = self.company.terminal_growth_rate
        terminal_growth_std = 0.005  # 终值增长率标准差

        for _ in range(iterations):
            # 随机采样
            sample_growth = np.random.normal(growth_rate_mean, growth_rate_std)
            sample_growth = max(0, sample_growth)  # 确保非负

            sample_margin = np.random.normal(margin_mean, margin_std)
            sample_margin = max(0.01, min(0.8, sample_margin))  # 限制在合理范围

            sample_wacc = np.random.normal(wacc_mean, wacc_std)
            sample_wacc = max(0.02, sample_wacc)  # 最小2%

            sample_terminal = np.random.normal(terminal_growth_mean, terminal_growth_std)
            sample_terminal = max(0, min(0.05, sample_terminal))  # 限制在0-5%

            # 计算估值
            custom_assumptions = {
                'growth_rate': sample_growth,
                'operating_margin': sample_margin,
            }

            try:
                valuation = AbsoluteValuation.dcf_valuation(
                    self.company,
                    wacc=sample_wacc,
                    terminal_growth_rate=sample_terminal,
                    custom_assumptions=custom_assumptions
                )
                values.append(valuation.value)
            except:
                # 如果计算失败，跳过本次迭代
                continue

        return MonteCarloResult(
            iterations=iterations,
            values=values
        )

    def generate_stress_report(self) -> Dict[str, Any]:
        """
        生成综合压力测试报告

        Returns:
            包含所有压力测试结果的报告
        """
        report = {
            'company': self.company.name,
            'base_valuation': self.base_valuation.value,
            'tests': {}
        }

        # 执行各类测试
        report['tests']['revenue_shock'] = [r.to_dict() for r in self.revenue_shock_test()]
        report['tests']['margin_compression'] = [r.to_dict() for r in self.margin_compression_test()]
        report['tests']['wacc_shock'] = [r.to_dict() for r in self.wacc_shock_test()]
        report['tests']['growth_slowdown'] = [r.to_dict() for r in self.growth_slowdown_test()]
        report['tests']['extreme_crash'] = self.extreme_market_crash().to_dict()

        # 蒙特卡洛模拟
        mc_result = self.monte_carlo_simulation(iterations=1000)
        report['monte_carlo'] = mc_result.to_dict()

        # 计算最大下行风险
        all_downsides = []
        for test_results in report['tests'].values():
            if isinstance(test_results, list):
                for r in test_results:
                    if r['change_pct'] < 0:
                        all_downsides.append(r['change_pct'])
            elif isinstance(test_results, dict):
                if test_results['change_pct'] < 0:
                    all_downsides.append(test_results['change_pct'])

        if all_downsides:
            report['max_downside'] = min(all_downsides)
        else:
            report['max_downside'] = 0

        return report


# ===== 辅助函数 =====

def display_stress_report(report: Dict[str, Any]) -> str:
    """
    格式化显示压力测试报告

    Args:
        report: 压力测试报告

    Returns:
        格式化的报告文本
    """
    output = []
    output.append("=" * 70)
    output.append(f"{report['company']} - 压力测试报告".center(68))
    output.append("=" * 70)

    base_valuation = report['base_valuation']
    output.append(f"\n基准估值: {base_valuation/10000:.2f}亿元\n")

    # 各项测试结果
    for test_name, test_results in report['tests'].items():
        output.append(f"【{test_name}】")

        if isinstance(test_results, list):
            for result in test_results:
                output.append(f"  {result['scenario_description']}")
                output.append(f"    估值: {result['stressed_value']/10000:.2f}亿元 "
                            f"({result['change_pct']:+.1%})")
        elif isinstance(test_results, dict):
            output.append(f"  {test_results['scenario_description']}")
            output.append(f"  估值: {test_results['stressed_value']/10000:.2f}亿元 "
                        f"({test_results['change_pct']:+.1%})")

        output.append("")

    # 蒙特卡洛结果
    if 'monte_carlo' in report:
        mc = report['monte_carlo']
        output.append("【蒙特卡洛模拟】")
        output.append(f"  迭代次数: {mc['iterations']}")
        output.append(f"  均值: {mc['mean']/10000:.2f}亿元")
        output.append(f"  中位数: {mc['median']/10000:.2f}亿元")
        output.append(f"  标准差: {mc['std']/10000:.2f}亿元")
        output.append(f"  范围: {mc['min_value']/10000:.2f} - {mc['max_value']/10000:.2f}亿元")
        output.append(f"  90%置信区间: {mc['confidence_interval_90'][0]/10000:.2f} - "
                    f"{mc['confidence_interval_90'][1]/10000:.2f}亿元")
        output.append("")

    # 最大下行风险
    output.append(f"【最大下行风险】{report['max_downside']:.1%}\n")

    output.append("=" * 70)

    return "\n".join(output)


# ===== 使用示例 =====

if __name__ == "__main__":
    from models import Company, CompanyStage

    # 创建测试公司
    company = Company(
        name="测试股份有限公司",
        industry="软件服务",
        stage=CompanyStage.GROWTH,
        revenue=50000,
        net_income=8000,
        ebitda=12000,
        net_assets=20000,
        total_debt=5000,
        cash_and_equivalents=2000,
        growth_rate=0.25,
        operating_margin=0.25,
        beta=1.2,
        terminal_growth_rate=0.025,
    )

    # 创建压力测试器
    tester = StressTester(company)

    # 生成综合报告
    report = tester.generate_stress_report()

    # 显示报告
    print(display_stress_report(report))

    # 单独运行蒙特卡洛模拟（更多迭代）
    print("\n【蒙特卡洛模拟详情】（5000次迭代）")
    mc_result = tester.monte_carlo_simulation(iterations=5000, seed=42)
    print(f"  均值: {mc_result.mean/10000:.2f}亿元")
    print(f"  25%分位: {mc_result.percentiles['p25']/10000:.2f}亿元")
    print(f"  75%分位: {mc_result.percentiles['p75']/10000:.2f}亿元")
    print(f"  90%置信区间: {mc_result.confidence_interval_90[0]/10000:.2f} - "
          f"{mc_result.confidence_interval_90[1]/10000:.2f}亿元")
