"""
投资估值系统使用示例
包含各种估值方法、情景分析、压力测试、敏感性分析的完整示例
"""
from models import Company, Comparable, CompanyStage, ValuationResult, ScenarioConfig
from relative_valuation import RelativeValuation
from absolute_valuation import AbsoluteValuation, display_dcf_details
from other_methods import OtherValuationMethods
from scenario_analysis import ScenarioAnalyzer, create_scenario_report
from stress_test import StressTester, display_stress_report
from sensitivity_analysis import SensitivityAnalyzer, display_sensitivity_results


def create_sample_company():
    """创建示例公司"""
    return Company(
        name="云数科技有限公司",
        industry="软件服务",
        stage=CompanyStage.GROWTH,
        revenue=50000,  # 5亿元
        net_income=8000,  # 8000万元
        ebitda=12000,
        gross_profit=32500,
        total_assets=30000,
        net_assets=20000,
        total_debt=5000,
        cash_and_equivalents=2000,
        growth_rate=0.25,
        margin=0.65,
        operating_margin=0.25,
        tax_rate=0.15,  # 高新技术企业优惠税率
        beta=1.2,
        risk_free_rate=0.03,
        market_risk_premium=0.07,
        cost_of_debt=0.05,
        target_debt_ratio=0.3,
        terminal_growth_rate=0.025,
        year=2024,
        description="国内领先的企业级云计算和大数据解决方案提供商",
    )


def create_sample_comparables():
    """创建可比公司示例"""
    return [
        Comparable(
            name="金山云",
            ts_code="3896.HK",
            industry="软件服务",
            revenue=80000,
            net_income=5000,
            net_assets=35000,
            ebitda=10000,
            pe_ratio=30.0,
            ps_ratio=5.5,
            pb_ratio=3.8,
            ev_ebitda=22.0,
            growth_rate=0.30,
        ),
        Comparable(
            name="用友网络",
            ts_code="600588.SH",
            revenue=95000,
            net_income=12000,
            net_assets=45000,
            ebitda=18000,
            pe_ratio=45.0,
            ps_ratio=6.8,
            pb_ratio=5.2,
            ev_ebitda=28.0,
            growth_rate=0.20,
        ),
        Comparable(
            name="恒生电子",
            ts_code="600570.SH",
            revenue=70000,
            net_income=15000,
            net_assets=40000,
            ebitda=20000,
            pe_ratio=35.0,
            ps_ratio=7.5,
            pb_ratio=4.5,
            ev_ebitda=25.0,
            growth_rate=0.18,
        ),
        Comparable(
            name="石基信息",
            ts_code="002153.SZ",
            revenue=40000,
            net_income=6000,
            net_assets=25000,
            ebitda=8000,
            pe_ratio=28.0,
            ps_ratio=5.0,
            pb_ratio=3.5,
            ev_ebitda=20.0,
            growth_rate=0.22,
        ),
    ]


def example_relative_valuation():
    """示例1：相对估值法"""
    print("=" * 70)
    print("示例1：相对估值法（P/E、P/S、P/B、EV/EBITDA）")
    print("=" * 70)

    company = create_sample_company()
    comparables = create_sample_comparables()

    print(f"\n目标公司: {company.name}")
    print(f"  行业: {company.industry}")
    print(f"  阶段: {company.stage.value}")
    print(f"  营业收入: {company.revenue/10000:.2f}亿元")
    print(f"  净利润: {company.net_income/10000:.2f}亿元")
    print(f"  净资产: {company.net_assets/10000:.2f}亿元")
    print(f"  预期增长率: {company.growth_rate:.1%}")

    print(f"\n可比公司数量: {len(comparables)}")
    for comp in comparables:
        print(f"  - {comp.name}: P/E={comp.pe_ratio:.1f}, P/S={comp.ps_ratio:.1f}")

    # 自动可比公司分析
    print("\n执行估值分析...")
    results = RelativeValuation.auto_comparable_analysis(company, comparables)

    print("\n" + "-" * 70)
    print("估值结果:")
    print("-" * 70)

    for method, result in results.items():
        print(f"\n{result}")
        if result.details:
            if 'pe_mean' in result.details:
                print(f"  可比P/E均值: {result.details['pe_mean']:.2f}")
            if 'ps_mean' in result.details:
                print(f"  可比P/S均值: {result.details['ps_mean']:.2f}")

    # 推荐估值区间
    if '综合' in results:
        final = results['综合']
        print(f"\n{'='*70}")
        print(f"推荐估值区间: {final.value_low/10000:.2f} - {final.value_high/10000:.2f}亿元")
        print(f"中值估值: {final.value_mid/10000:.2f}亿元")
        print(f"{'='*70}")


def example_dcf_valuation():
    """示例2：DCF绝对估值法"""
    print("\n\n" + "=" * 70)
    print("示例2：现金流折现模型（DCF）")
    print("=" * 70)

    company = create_sample_company()

    # DCF估值
    result = AbsoluteValuation.dcf_valuation(
        company,
        projection_years=5,
    )

    print(display_dcf_details(result))


def example_scenario_analysis():
    """示例3：情景分析"""
    print("\n\n" + "=" * 70)
    print("示例3：情景分析（基准/乐观/悲观）")
    print("=" * 70)

    company = create_sample_company()
    analyzer = ScenarioAnalyzer(company)

    # 三情景分析
    results = analyzer.compare_scenarios()

    # 生成报告
    report = create_scenario_report(results, format="text")
    print(report)


def example_stress_test():
    """示例4：压力测试"""
    print("\n\n" + "=" * 70)
    print("示例4：压力测试")
    print("=" * 70)

    company = create_sample_company()
    tester = StressTester(company)

    # 生成综合报告
    report = tester.generate_stress_report()

    # 显示报告
    print(display_stress_report(report))


def example_sensitivity_analysis():
    """示例5：敏感性分析"""
    print("\n\n" + "=" * 70)
    print("示例5：敏感性分析")
    print("=" * 70)

    company = create_sample_company()
    analyzer = SensitivityAnalyzer(company)

    # 综合敏感性分析
    results = analyzer.comprehensive_sensitivity_analysis()

    # 显示结果
    print(display_sensitivity_results(results))


def example_vc_method():
    """示例6：VC法（早期项目）"""
    print("\n\n" + "=" * 70)
    print("示例6：风险投资法（早期项目）")
    print("=" * 70)

    # 早期科技公司 - 刚开始盈利
    early_company = Company(
        name="AI创新科技（早期项目）",
        industry="人工智能",
        stage=CompanyStage.EARLY,
        revenue=3000,  # 3000万
        net_income=200,  # 刚开始盈利，200万
        net_assets=4000,
        growth_rate=0.60,  # 高增长但更合理
        operating_margin=0.08,  # 利润率较低
    )

    print(f"\n目标公司: {early_company.name}")
    print(f"  阶段: {early_company.stage.value}")
    print(f"  收入: {early_company.revenue/10000:.2f}亿元")
    print(f"  净利润: {early_company.net_income/10000:.2f}亿元")
    print(f"  预期增长率: {early_company.growth_rate:.0%}")

    # VC法估值
    result = OtherValuationMethods.vc_method_with_future_projection(
        early_company,
        projection_years=5,
        target_pe=35.0,
        target_return_multiple=12.0,
        margin_improvement=0.03,
    )

    print(f"\n{result}")
    print(f"\n假设:")
    print(f"  5年后净利润: {result.details['future_net_income']/10000:.2f}亿元")
    print(f"  退出时P/E: {result.details['target_pe']:.1f}x")
    print(f"  退出估值: {result.details['exit_valuation']/10000:.2f}亿元")
    print(f"  目标回报倍数: {result.details['target_return_multiple']:.0f}x")
    print(f"  当前投后估值: {result.value/10000:.2f}亿元")
    if 'implied_irr' in result.details:
        print(f"  隐含IRR: {result.details['implied_irr']:.2%}")


def example_comprehensive_valuation():
    """示例7：综合估值流程"""
    print("\n\n" + "=" * 70)
    print("示例7：综合估值分析")
    print("=" * 70)

    company = create_sample_company()
    comparables = create_sample_comparables()

    print(f"\n【公司概况】")
    print(f"  名称: {company.name}")
    print(f"  行业: {company.industry}")
    print(f"  阶段: {company.stage.value}")
    print(f"  收入: {company.revenue/10000:.2f}亿元")
    print(f"  净利润: {company.net_income/10000:.2f}亿元")
    print(f"  增长率: {company.growth_rate:.1%}")

    # 1. 相对估值
    print(f"\n{'='*70}")
    print("【第一步：相对估值】")
    print('='*70)
    relative_results = RelativeValuation.auto_comparable_analysis(company, comparables)
    for method, result in relative_results.items():
        print(f"  {method}: {result.value/10000:.2f}亿元")

    # 2. 绝对估值
    print(f"\n{'='*70}")
    print("【第二步：绝对估值（DCF）】")
    print('='*70)
    dcf_result = AbsoluteValuation.dcf_valuation(company)
    print(f"  DCF估值: {dcf_result.value/10000:.2f}亿元")

    # 3. 情景分析
    print(f"\n{'='*70}")
    print("【第三步：情景分析】")
    print('='*70)
    scenario_analyzer = ScenarioAnalyzer(company)
    scenario_results = scenario_analyzer.compare_scenarios()

    for name, data in scenario_results.items():
        if name == 'statistics':
            continue
        print(f"  {name}: {data['value']/10000:.2f}亿元")

    # 4. 压力测试
    print(f"\n{'='*70}")
    print("【第四步：压力测试】")
    print('='*70)
    stress_tester = StressTester(company)
    extreme_result = stress_tester.extreme_market_crash()
    print(f"  极端情景估值: {extreme_result.stressed_value/10000:.2f}亿元 "
          f"({extreme_result.change_pct:+.1%})")

    # 5. 综合估值建议
    print(f"\n{'='*70}")
    print("【综合估值建议】")
    print('='*70)

    all_values = []
    if '综合' in relative_results:
        all_values.append(relative_results['综合'].value)
    all_values.append(dcf_result.value)

    if 'statistics' in scenario_results:
        all_values.append(scenario_results['statistics']['median'])

    import numpy as np
    final_value = np.median(all_values)

    print(f"\n  估值方法结果:")
    print(f"    相对估值（综合）: {relative_results.get('综合', {}).value/10000 if '综合' in relative_results else 'N/A'}亿元")
    print(f"    DCF估值: {dcf_result.value/10000:.2f}亿元")
    print(f"    情景分析中值: {scenario_results.get('statistics', {}).get('median', 0)/10000:.2f}亿元")

    print(f"\n  【推荐估值】: {final_value/10000:.2f}亿元")

    # 估值区间
    value_range = (
        min(all_values) * 0.9,
        max(all_values) * 1.1
    )
    print(f"  估值区间: {value_range[0]/10000:.2f} - {value_range[1]/10000:.2f}亿元")

    # 风险提示
    print(f"\n  风险提示:")
    print(f"    极端情景下估值: {extreme_result.stressed_value/10000:.2f}亿元")
    print(f"    最大下行风险: {extreme_result.change_pct:.1%}")

    print(f"\n{'='*70}")


def run_all_examples():
    """运行所有示例"""
    print("\n")
    print("#" * 70)
    print("#" + " " * 68 + "#")
    print("#" + "  投资估值系统 - 完整示例演示  ".center(68) + "#")
    print("#" + " " * 68 + "#")
    print("#" * 70)

    example_relative_valuation()
    example_dcf_valuation()
    example_scenario_analysis()
    example_stress_test()
    example_sensitivity_analysis()
    example_vc_method()
    example_comprehensive_valuation()

    print("\n\n" + "#" * 70)
    print("#" + " " * 68 + "#")
    print("#" + "  示例演示完成  ".center(68) + "#")
    print("#" + " " * 68 + "#")
    print("#" * 70)
    print()


if __name__ == "__main__":
    # 运行所有示例
    run_all_examples()

    # 或者单独运行某个示例
    # example_relative_valuation()
    # example_dcf_valuation()
    # example_scenario_analysis()
    # example_stress_test()
    # example_sensitivity_analysis()
    # example_vc_method()
    # example_comprehensive_valuation()
