"""
绝对估值法模块 - 现金流折现模型（DCF）
包含自由现金流预测、WACC计算、终值计算等
"""
import numpy as np
from typing import List, Optional, Dict, Any, Tuple
from core.models import Company, ValuationResult


class AbsoluteValuation:
    """绝对估值（DCF）计算类"""

    @staticmethod
    def calculate_wacc(
        company: Company,
        use_capm: bool = True,
        market_risk_premium: Optional[float] = None,
        beta: Optional[float] = None
    ) -> float:
        """
        计算加权平均资本成本（WACC）

        Args:
            company: 公司对象
            use_capm: 是否使用CAPM计算股权成本
            market_risk_premium: 市场风险溢价（如提供则覆盖公司参数）
            beta: 贝塔系数（如提供则覆盖公司参数）

        Returns:
            WACC（小数形式，如0.08表示8%）
        """
        # 股权成本（CAPM模型）
        if use_capm:
            beta_val = beta if beta is not None else company.beta
            mrp = market_risk_premium if market_risk_premium is not None else company.market_risk_premium
            cost_of_equity = company.risk_free_rate + beta_val * mrp
        else:
            # 简化处理：假设股权成本 = 债务成本 + 股权风险溢价
            cost_of_equity = company.cost_of_debt + 0.05

        # 债务成本
        cost_of_debt = company.cost_of_debt * (1 - company.tax_rate)  # 税后债务成本

        # 目标资本结构
        debt_ratio = company.target_debt_ratio
        equity_ratio = 1 - debt_ratio

        # WACC计算
        wacc = cost_of_equity * equity_ratio + cost_of_debt * debt_ratio

        return wacc

    @staticmethod
    def calculate_terminal_value(
        final_fcf: float,
        wacc: float,
        terminal_growth_rate: float,
        method: str = "perpetuity"
    ) -> float:
        """
        计算终值

        Args:
            final_fcf: 预测期最后一年的自由现金流
            wacc: 加权平均资本成本
            terminal_growth_rate: 永续增长率
            method: 计算方法，"perpetuity"（永续增长）或 "exit_multiple"（退出倍数）

        Returns:
            终值（预测期末的现值）
        """
        if method == "perpetuity":
            # 永续增长模型（戈登模型）
            if wacc <= terminal_growth_rate:
                raise ValueError(f"WACC({wacc:.2%})必须大于永续增长率({terminal_growth_rate:.2%})")
            terminal_value = final_fcf * (1 + terminal_growth_rate) / (wacc - terminal_growth_rate)
        else:
            # 退出倍数法（简化：假设FCF的某个倍数）
            # 实际应用中应该使用EBITDA退出倍数
            multiple = 10.0  # 默认10倍
            terminal_value = final_fcf * multiple

        return terminal_value

    @staticmethod
    def forecast_free_cash_flows(
        company: Company,
        projection_years: int = 5,
        custom_assumptions: Optional[Dict[str, Any]] = None
    ) -> List[Dict[str, float]]:
        """
        预测自由现金流

        Args:
            company: 公司对象
            projection_years: 预测年数
            custom_assumptions: 自定义假设（可覆盖公司默认参数）

        Returns:
            包含每年预测数据的字典列表
        """
        forecasts = []

        # 使用自定义假设或公司默认参数
        growth_rate = custom_assumptions.get('growth_rate', company.growth_rate) if custom_assumptions else company.growth_rate
        margin = custom_assumptions.get('margin', company.margin) if custom_assumptions else company.margin
        operating_margin = custom_assumptions.get('operating_margin', company.operating_margin) if custom_assumptions else company.operating_margin

        # 如果没有营业利润率，用毛利率估算
        if not operating_margin and margin:
            operating_margin = margin * 0.3  # 假设营业利润约为毛利的30%

        # 假设资本支出和营运资金变化占收入的比例
        capex_ratio = custom_assumptions.get('capex_ratio', 0.05) if custom_assumptions else 0.05
        wc_change_ratio = custom_assumptions.get('wc_change_ratio', 0.02) if custom_assumptions else 0.02
        depreciation_ratio = custom_assumptions.get('depreciation_ratio', 0.03) if custom_assumptions else 0.03

        # 初始收入（使用局部变量，避免修改原对象）
        revenue = company.revenue

        for year in range(1, projection_years + 1):
            # 收入增长（可设置逐年递减）
            if year > 3:
                year_growth = max(growth_rate * 0.7, 0.05)  # 后期增长率放缓
            else:
                year_growth = growth_rate

            revenue = revenue * (1 + year_growth)

            # 利润计算
            operating_profit = revenue * operating_margin if operating_margin else 0
            tax = operating_profit * company.tax_rate
            nopat = operating_profit - tax  # 税后营业利润

            # 折旧摊销（简化：占收入比例）
            depreciation = revenue * depreciation_ratio

            # 自由现金流 = NOPAT + 折旧 - 资本支出 - 营运资金增加
            capex = revenue * capex_ratio
            wc_change = revenue * wc_change_ratio

            fcf = nopat + depreciation - capex - wc_change

            forecasts.append({
                'year': year,
                'revenue': revenue,
                'operating_profit': operating_profit,
                'nopat': nopat,
                'depreciation': depreciation,
                'capex': capex,
                'wc_change': wc_change,
                'fcf': fcf,
                'growth_rate': year_growth,
            })

        return forecasts

    @staticmethod
    def dcf_valuation(
        company: Company,
        projection_years: int = 5,
        wacc: Optional[float] = None,
        terminal_growth_rate: Optional[float] = None,
        terminal_method: str = "perpetuity",
        custom_assumptions: Optional[Dict[str, Any]] = None
    ) -> ValuationResult:
        """
        现金流折现模型估值

        Args:
            company: 公司对象
            projection_years: 预测年数
            wacc: 加权平均资本成本（如不提供则自动计算）
            terminal_growth_rate: 永续增长率（如不提供则使用公司参数）
            terminal_method: 终值计算方法
            custom_assumptions: 自定义假设

        Returns:
            DCF估值结果
        """
        # 计算WACC
        if wacc is None:
            wacc = AbsoluteValuation.calculate_wacc(company)

        # 获取永续增长率
        if terminal_growth_rate is None:
            terminal_growth_rate = company.terminal_growth_rate

        # 预测自由现金流
        fcf_forecasts = AbsoluteValuation.forecast_free_cash_flows(
            company, projection_years, custom_assumptions
        )

        # 折现计算现值
        pv_forecasts = 0
        for forecast in fcf_forecasts:
            year = forecast['year']
            fcf = forecast['fcf']
            pv = fcf / ((1 + wacc) ** year)
            pv_forecasts += pv

        # 计算终值
        final_fcf = fcf_forecasts[-1]['fcf']
        terminal_value = AbsoluteValuation.calculate_terminal_value(
            final_fcf, wacc, terminal_growth_rate, terminal_method
        )

        # 终值折现
        pv_terminal = terminal_value / ((1 + wacc) ** projection_years)

        # 企业价值 = 预测期现值 + 终值现值
        enterprise_value = pv_forecasts + pv_terminal

        # 股权价值 = 企业价值 - 净债务
        equity_value = enterprise_value - company.net_debt

        return ValuationResult(
            method="DCF法",
            value=equity_value,
            details={
                'wacc': wacc,
                'projection_years': projection_years,
                'terminal_growth_rate': terminal_growth_rate,
                'terminal_method': terminal_method,
                'pv_forecasts': pv_forecasts,
                'pv_terminal': pv_terminal,
                'enterprise_value': enterprise_value,
                'terminal_value': terminal_value,
                'net_debt': company.net_debt,
                'fcf_forecasts': fcf_forecasts,
            },
            assumptions=custom_assumptions or {
                'growth_rate': company.growth_rate,
                'operating_margin': company.operating_margin,
                'capex_ratio': 0.05,
                'wc_change_ratio': 0.02,
                'depreciation_ratio': 0.03,
            }
        )

    @staticmethod
    def dcf_sensitivity_analysis(
        company: Company,
        wacc_range: Tuple[float, float] = (0.06, 0.12),
        growth_rate_range: Tuple[float, float] = (0.05, 0.25),
        terminal_growth_range: Tuple[float, float] = (0.01, 0.04),
        steps: int = 5
    ) -> Dict[str, Any]:
        """
        DCF敏感性分析

        Args:
            company: 公司对象
            wacc_range: WACC范围
            growth_rate_range: 增长率范围
            terminal_growth_range: 永续增长率范围
            steps: 每个参数的步数

        Returns:
            敏感性分析结果
        """
        results = {}

        # WACC敏感性
        wacc_values = np.linspace(wacc_range[0], wacc_range[1], steps)
        results['wacc_sensitivity'] = []
        for wacc in wacc_values:
            valuation = AbsoluteValuation.dcf_valuation(company, wacc=wacc)
            results['wacc_sensitivity'].append({
                'parameter': 'wacc',
                'value': wacc,
                'valuation': valuation.value,
            })

        # 增长率敏感性
        growth_values = np.linspace(growth_rate_range[0], growth_rate_range[1], steps)
        results['growth_sensitivity'] = []
        for growth in growth_values:
            assumptions = {'growth_rate': growth}
            valuation = AbsoluteValuation.dcf_valuation(company, custom_assumptions=assumptions)
            results['growth_sensitivity'].append({
                'parameter': 'growth_rate',
                'value': growth,
                'valuation': valuation.value,
            })

        # 永续增长率敏感性
        terminal_values = np.linspace(terminal_growth_range[0], terminal_growth_range[1], steps)
        results['terminal_growth_sensitivity'] = []
        for terminal_growth in terminal_values:
            valuation = AbsoluteValuation.dcf_valuation(company, terminal_growth_rate=terminal_growth)
            results['terminal_growth_sensitivity'].append({
                'parameter': 'terminal_growth',
                'value': terminal_growth,
                'valuation': valuation.value,
            })

        # 计算敏感度（估值变化百分比）
        for key in ['wacc_sensitivity', 'growth_sensitivity', 'terminal_growth_sensitivity']:
            data = results[key]
            if len(data) >= 2:
                base_valuation = data[len(data) // 2]['valuation']
                min_valuation = min(d['valuation'] for d in data)
                max_valuation = max(d['valuation'] for d in data)
                results[f'{key}_stats'] = {
                    'base': base_valuation,
                    'min': min_valuation,
                    'max': max_valuation,
                    'range_pct': (max_valuation - min_valuation) / base_valuation if base_valuation > 0 else 0,
                }

        return results


# ===== 辅助函数 =====

def calculate_fcf_from_income(
    revenue: float,
    ebit: float,
    tax_rate: float,
    depreciation: float,
    capex: float,
    wc_change: float
) -> float:
    """
    从利润表数据计算自由现金流

    FCF = EBIT(1-t) + 折旧 - 资本支出 - 营运资金变化
    """
    nopat = ebit * (1 - tax_rate)
    return nopat + depreciation - capex - wc_change


def display_dcf_details(result: ValuationResult) -> str:
    """
    格式化显示DCF估值详情
    """
    output = []
    output.append("=" * 60)
    output.append(f"{result.method}估值详情")
    output.append("=" * 60)

    details = result.details

    output.append(f"\n【基本参数】")
    output.append(f"  WACC: {details['wacc']:.2%}")
    output.append(f"  预测期: {details['projection_years']}年")
    output.append(f"  永续增长率: {details['terminal_growth_rate']:.2%}")
    output.append(f"  终值方法: {details['terminal_method']}")

    output.append(f"\n【预测期现金流】")
    for fcf in details['fcf_forecasts']:
        output.append(f"  第{fcf['year']}年: 收入{fcf['revenue']/10000:.1f}亿, "
                    f"FCF={fcf['fcf']/10000:.2f}亿")

    output.append(f"\n【估值构成】")
    output.append(f"  预测期现值: {details['pv_forecasts']/10000:.2f}亿元")
    output.append(f"  终值: {details['terminal_value']/10000:.2f}亿元")
    output.append(f"  终值现值: {details['pv_terminal']/10000:.2f}亿元")
    output.append(f"  企业价值: {details['enterprise_value']/10000:.2f}亿元")
    output.append(f"  净债务: {details['net_debt']/10000:.2f}亿元")
    output.append(f"  股权价值: {result.value/10000:.2f}亿元")

    output.append("\n" + "=" * 60)

    return "\n".join(output)


# ===== 使用示例 =====

if __name__ == "__main__":
    from models import Company, CompanyStage

    # 创建测试公司
    company = Company(
        name="测试股份有限公司",
        industry="软件服务",
        stage=CompanyStage.GROWTH,
        revenue=50000,  # 5亿元
        net_income=8000,  # 8000万元
        ebitda=12000,
        total_assets=30000,
        net_assets=20000,
        total_debt=5000,
        cash_and_equivalents=2000,
        growth_rate=0.20,
        operating_margin=0.25,
        beta=1.2,
        risk_free_rate=0.03,
        market_risk_premium=0.07,
        cost_of_debt=0.05,
        target_debt_ratio=0.3,
        terminal_growth_rate=0.025,
    )

    # DCF估值
    result = AbsoluteValuation.dcf_valuation(
        company,
        projection_years=5,
    )

    print(display_dcf_details(result))

    # 敏感性分析
    print("\n【敏感性分析】")
    sensitivity = AbsoluteValuation.dcf_sensitivity_analysis(company)

    print("\nWACC敏感性:")
    for item in sensitivity['wacc_sensitivity']:
        print(f"  WACC={item['value']:.2%}: {item['valuation']/10000:.2f}亿元")

    print("\n增长率敏感性:")
    for item in sensitivity['growth_sensitivity']:
        print(f"  增长率={item['value']:.2%}: {item['valuation']/10000:.2f}亿元")
