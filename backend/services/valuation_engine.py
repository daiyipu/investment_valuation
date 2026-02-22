"""
统一估值引擎
提供统一的估值入口，整合多种估值方法，自动生成交叉验证和综合报告
"""
from typing import Optional, List, Dict, Any, Tuple
from datetime import datetime

from core.models import Company, Comparable, CompanyStage, ValuationResult, ScenarioConfig
from services.relative_valuation import RelativeValuation
from services.absolute_valuation import AbsoluteValuation
from utils.other_methods import OtherValuationMethods
from services.scenario_analysis import ScenarioAnalyzer, SCENARIOS
from services.stress_test import StressTester
from services.sensitivity_analysis import SensitivityAnalyzer
from data_fetcher import TushareDataFetcher


class ValuationEngine:
    """
    统一估值引擎

    提供一站式估值服务，整合所有估值方法和风险分析功能
    """

    def __init__(self, tushare_token: Optional[str] = None):
        """
        初始化估值引擎

        Args:
            tushare_token: Tushare API Token，用于获取可比公司数据
        """
        self.tushare_token = tushare_token
        self.fetcher = TushareDataFetcher(tushare_token) if tushare_token else None

    def full_valuation(
        self,
        company: Company,
        comparables: Optional[List[Comparable]] = None,
        methods: Optional[List[str]] = None,
        enable_risk_analysis: bool = True,
        industry_for_comparables: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        完整估值流程

        执行多种估值方法、情景分析、压力测试、敏感性分析，生成综合报告

        Args:
            company: 目标公司
            comparables: 可比公司列表（如不提供且配置了tushare_token则自动获取）
            methods: 要使用的估值方法列表
            enable_risk_analysis: 是否执行风险分析
            industry_for_comparables: 行业名称（用于自动获取可比公司）

        Returns:
            综合估值报告
        """
        report = {
            'company': company.name,
            'industry': company.industry,
            'stage': company.stage.value,
            'timestamp': datetime.now().isoformat(),
            'valuation_methods': {},
            'risk_analysis': {},
            'recommendation': {}
        }

        # ========== 第一步：相对估值 ==========
        if comparables is None and self.fetcher and industry_for_comparables:
            print(f"正在从Tushare获取{industry_for_comparables}行业可比公司...")
            comparables = self.fetcher.get_comparable_companies(industry_for_comparables)

        if comparables:
            print(f"使用{len(comparables)}家可比公司进行相对估值...")
            relative_results = RelativeValuation.auto_comparable_analysis(
                company, comparables, methods
            )
            report['valuation_methods']['relative'] = {
                name: result.to_dict() for name, result in relative_results.items()
            }

        # ========== 第二步：绝对估值（DCF）==========
        print("执行DCF估值...")
        try:
            dcf_result = AbsoluteValuation.dcf_valuation(company)
            report['valuation_methods']['absolute'] = {
                'DCF': dcf_result.to_dict()
            }
        except Exception as e:
            print(f"DCF估值失败: {e}")

        # ========== 第三步：风险分析 ==========
        if enable_risk_analysis:
            print("执行风险分析...")

            # 情景分析
            try:
                scenario_analyzer = ScenarioAnalyzer(company, dcf_result)
                scenario_results = scenario_analyzer.compare_scenarios()
                report['risk_analysis']['scenario'] = scenario_results
            except Exception as e:
                print(f"情景分析失败: {e}")

            # 压力测试
            try:
                stress_tester = StressTester(company)
                stress_report = stress_tester.generate_stress_report()
                report['risk_analysis']['stress_test'] = stress_report
            except Exception as e:
                print(f"压力测试失败: {e}")

            # 敏感性分析
            try:
                sensitivity_analyzer = SensitivityAnalyzer(company)
                sensitivity_results = sensitivity_analyzer.comprehensive_sensitivity_analysis()
                report['risk_analysis']['sensitivity'] = sensitivity_results
            except Exception as e:
                print(f"敏感性分析失败: {e}")

        # ========== 第四步：交叉验证与建议 ==========
        all_values = []
        method_details = []

        # 收集所有估值结果
        if 'relative' in report['valuation_methods']:
            for method, result in report['valuation_methods']['relative'].items():
                if result['value'] and result['value'] > 0:
                    all_values.append(result['value'])
                    method_details.append({
                        'method': f"相对估值-{method}",
                        'value': result['value'],
                        'range': (result.get('value_low'), result.get('value_high'))
                    })

        if 'absolute' in report['valuation_methods']:
            for method, result in report['valuation_methods']['absolute'].items():
                if result['value'] and result['value'] > 0:
                    all_values.append(result['value'])
                    method_details.append({
                        'method': f"绝对估值-{method}",
                        'value': result['value'],
                    })

        # 计算推荐估值
        if all_values:
            import numpy as np
            final_value = np.median(all_values)
            value_min = min(all_values) * 0.9
            value_max = max(all_values) * 1.1

            report['recommendation'] = {
                'final_value': final_value,
                'value_range': (value_min, value_max),
                'confidence': self._calculate_confidence(all_values),
                'methods_used': len(all_values),
                'method_details': method_details,
            }

        # ========== 第五步：生成文本报告 ==========
        report['text_report'] = self._generate_text_report(report)

        return report

    def quick_valuation(
        self,
        company: Company,
        method: str = 'auto'
    ) -> ValuationResult:
        """
        快速估值

        根据公司阶段自动选择最合适的估值方法

        Args:
            company: 目标公司
            method: 指定估值方法（'auto'为自动选择）

        Returns:
            估值结果
        """
        stage = company.stage

        if method == 'auto':
            # 根据公司阶段自动选择方法
            if stage == CompanyStage.EARLY:
                # 早期项目：优先使用P/S法或VC法
                if company.revenue > 0 and company.net_income <= 0:
                    method = 'PS'
                else:
                    method = 'VC'
            elif stage == CompanyStage.GROWTH:
                # 成长期：优先使用P/S法或DCF
                method = 'PS' if company.net_income <= 0 else 'DCF'
            else:
                # 成熟期/上市：优先使用P/E法或DCF
                method = 'PE' if company.net_income > 0 else 'DCF'

        # 执行估值
        if method == 'DCF':
            return AbsoluteValuation.dcf_valuation(company)
        elif method == 'PE':
            # 需要可比公司
            raise ValueError("PE法需要可比公司数据")
        elif method == 'PS':
            raise ValueError("PS法需要可比公司数据")
        elif method == 'VC':
            return OtherValuationMethods.vc_method_with_future_projection(company)
        else:
            raise ValueError(f"不支持的估值方法: {method}")

    def auto_fetch_and_valuate(
        self,
        company: Company,
        industry: str
    ) -> Dict[str, Any]:
        """
        自动获取可比公司并估值

        Args:
            company: 目标公司
            industry: 所属行业

        Returns:
            估值报告
        """
        if not self.fetcher:
            raise ValueError("需要配置Tushare Token才能使用此功能")

        # 获取可比公司
        comparables = self.fetcher.get_comparable_companies(industry)

        if not comparables:
            raise ValueError(f"未找到{industry}行业的可比公司")

        # 执行完整估值
        return self.full_valuation(
            company,
            comparables=comparables,
            industry_for_comparables=industry
        )

    def batch_valuation(
        self,
        companies: List[Company],
        comparables: Optional[List[Comparable]] = None
    ) -> List[Dict[str, Any]]:
        """
        批量估值

        对多个公司进行估值

        Args:
            companies: 公司列表
            comparables: 共用的可比公司列表

        Returns:
            估值结果列表
        """
        results = []

        for i, company in enumerate(companies):
            print(f"\n正在估值第{i+1}/{len(companies)}家公司: {company.name}")

            try:
                report = self.full_valuation(company, comparables, enable_risk_analysis=False)
                results.append({
                    'company': company.name,
                    'success': True,
                    'report': report
                })
            except Exception as e:
                results.append({
                    'company': company.name,
                    'success': False,
                    'error': str(e)
                })

        return results

    def compare_scenarios(
        self,
        company: Company,
        custom_scenarios: Optional[List[ScenarioConfig]] = None
    ) -> Dict[str, Any]:
        """
        多情景估值对比

        Args:
            company: 目标公司
            custom_scenarios: 自定义情景列表

        Returns:
            情景对比结果
        """
        analyzer = ScenarioAnalyzer(company)

        if custom_scenarios:
            results = analyzer.compare_scenarios(scenarios=custom_scenarios)
        else:
            results = analyzer.compare_scenarios()

        return results

    def generate_report(
        self,
        report_data: Dict[str, Any],
        format: str = 'text'
    ) -> str:
        """
        生成格式化报告

        Args:
            report_data: 估值报告数据
            format: 报告格式（text/markdown/html）

        Returns:
            格式化的报告文本
        """
        if format == 'text':
            return self._generate_text_report(report_data)
        elif format == 'markdown':
            return self._generate_markdown_report(report_data)
        elif format == 'html':
            return self._generate_html_report(report_data)
        else:
            raise ValueError(f"不支持的格式: {format}")

    # ===== 私有辅助方法 =====

    def _calculate_confidence(self, values: List[float]) -> str:
        """计算估值置信度"""
        import numpy as np
        std = np.std(values)
        mean = np.mean(values)
        cv = std / mean if mean > 0 else 1

        if cv < 0.1:
            return "高"
        elif cv < 0.2:
            return "中"
        else:
            return "低"

    def _generate_text_report(self, report: Dict[str, Any]) -> str:
        """生成文本格式报告"""
        lines = []
        lines.append("=" * 70)
        lines.append(f"{report['company']} - 估值报告".center(68))
        lines.append("=" * 70)
        lines.append(f"行业: {report['industry']}")
        lines.append(f"阶段: {report['stage']}")
        lines.append(f"时间: {report['timestamp']}")
        lines.append("")

        # 估值方法结果
        lines.append("-" * 70)
        lines.append("【估值方法】")
        lines.append("-" * 70)

        if 'relative' in report.get('valuation_methods', {}):
            for method, result in report['valuation_methods']['relative'].items():
                lines.append(f"\n{method}: {result['value']/10000:.2f}亿元")

        if 'absolute' in report.get('valuation_methods', {}):
            for method, result in report['valuation_methods']['absolute'].items():
                lines.append(f"\n{method}: {result['value']/10000:.2f}亿元")

        # 推荐估值
        if 'recommendation' in report and report['recommendation']:
            rec = report['recommendation']
            lines.append("\n" + "-" * 70)
            lines.append("【估值建议】")
            lines.append("-" * 70)
            lines.append(f"\n推荐估值: {rec['final_value']/10000:.2f}亿元")
            lines.append(f"估值区间: {rec['value_range'][0]/10000:.2f} - {rec['value_range'][1]/10000:.2f}亿元")
            lines.append(f"置信度: {rec.get('confidence', 'N/A')}")

        # 风险分析
        if 'risk_analysis' in report:
            lines.append("\n" + "-" * 70)
            lines.append("【风险分析】")
            lines.append("-" * 70)

            if 'stress_test' in report['risk_analysis']:
                stress = report['risk_analysis']['stress_test']
                if 'max_downside' in stress:
                    lines.append(f"\n最大下行风险: {stress['max_downside']:.1%}")

        lines.append("\n" + "=" * 70)

        return "\n".join(lines)

    def _generate_markdown_report(self, report: Dict[str, Any]) -> str:
        """生成Markdown格式报告"""
        lines = []
        lines.append(f"# {report['company']} - 估值报告\n")
        lines.append(f"- **行业**: {report['industry']}")
        lines.append(f"- **阶段**: {report['stage']}")
        lines.append(f"- **时间**: {report['timestamp']}\n")

        lines.append("## 估值方法\n")

        if 'recommendation' in report and report['recommendation']:
            rec = report['recommendation']
            lines.append(f"### 推荐估值\n")
            lines.append(f"- **估值**: {rec['final_value']/10000:.2f}亿元")
            lines.append(f"- **区间**: {rec['value_range'][0]/10000:.2f} - {rec['value_range'][1]/10000:.2f}亿元")
            lines.append(f"- **置信度**: {rec.get('confidence', 'N/A')}\n")

        if 'risk_analysis' in report:
            lines.append("## 风险分析\n")

            if 'scenario' in report['risk_analysis']:
                scenario = report['risk_analysis']['scenario']
                if 'statistics' in scenario:
                    stats = scenario['statistics']
                    lines.append(f"- **均值**: {stats['mean']/10000:.2f}亿元")
                    lines.append(f"- **标准差**: {stats['std']/10000:.2f}亿元")

        return "\n".join(lines)

    def _generate_html_report(self, report: Dict[str, Any]) -> str:
        """生成HTML格式报告"""
        html = f"""
        <html>
        <head>
            <title>{report['company']} - 估值报告</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 20px; }}
                .header {{ background: #667eea; color: white; padding: 20px; }}
                .section {{ margin: 20px 0; padding: 15px; border: 1px solid #ddd; }}
                .valuation {{ font-size: 24px; color: #667eea; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>{report['company']} - 估值报告</h1>
            </div>
            <div class="section">
                <p>行业: {report['industry']} | 阶段: {report['stage']}</p>
            </div>
            <div class="section">
                <h2>推荐估值</h2>
                <p class="valuation">{report['recommendation']['final_value']/10000:.2f} 亿元</p>
            </div>
        </body>
        </html>
        """
        return html


# ===== 使用示例 =====

if __name__ == "__main__":
    from models import Company, Comparable, CompanyStage

    # 创建测试公司
    company = Company(
        name="测试科技股份有限公司",
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

    # 创建可比公司
    comparables = [
        Comparable(
            name="可比公司A",
            revenue=80000,
            net_income=15000,
            net_assets=30000,
            pe_ratio=30.0,
            ps_ratio=6.0,
            pb_ratio=4.0,
        ),
        Comparable(
            name="可比公司B",
            revenue=60000,
            net_income=10000,
            net_assets=25000,
            pe_ratio=25.0,
            ps_ratio=5.0,
            pb_ratio=3.5,
        ),
    ]

    # 创建估值引擎
    engine = ValuationEngine()

    # 执行完整估值
    print("执行完整估值流程...")
    report = engine.full_valuation(company, comparables)

    # 打印文本报告
    print("\n" + report['text_report'])
