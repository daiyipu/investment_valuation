"""
多产品DCF估值模块
支持对公司多个产品/业务线分别估值，然后叠加得到公司整体估值
"""
from typing import List, Dict, Any, Optional, Tuple
from models import ProductSegment, ProductValuationResult, MultiProductValuationResult


class MultiProductValuation:
    """多产品估值类"""

    @staticmethod
    def calculate_wacc(
        beta: float,
        risk_free_rate: float = 0.03,
        market_risk_premium: float = 0.07,
        cost_of_debt: float = 0.05,
        target_debt_ratio: float = 0.3,
        tax_rate: float = 0.25
    ) -> float:
        """
        计算加权平均资本成本（WACC）

        Args:
            beta: 贝塔系数
            risk_free_rate: 无风险利率
            market_risk_premium: 市场风险溢价
            cost_of_debt: 债务成本
            target_debt_ratio: 目标资产负债率
            tax_rate: 所得税率

        Returns:
            WACC
        """
        # 股权成本 = 无风险利率 + β × 市场风险溢价
        cost_of_equity = risk_free_rate + beta * market_risk_premium

        # 税后债务成本
        after_tax_cost_of_debt = cost_of_debt * (1 - tax_rate)

        # WACC = 股权成本 × 股权占比 + 债务成本 × 债务占比
        equity_ratio = 1 - target_debt_ratio
        wacc = cost_of_equity * equity_ratio + after_tax_cost_of_debt * target_debt_ratio

        return wacc

    @staticmethod
    def forecast_product_cash_flows(
        product: ProductSegment,
        projection_years: int = 5,
        tax_rate: float = 0.25
    ) -> List[Dict[str, float]]:
        """
        预测单个产品的自由现金流

        Args:
            product: 产品对象
            projection_years: 预测年数
            tax_rate: 所得税率

        Returns:
            包含每年预测数据的字典列表
        """
        forecasts = []
        revenue = product.current_revenue

        for year in range(1, projection_years + 1):
            # 获取该年的增长率（如果产品增长率列表不够长，使用最后一个增长率）
            if year <= len(product.growth_rate_years):
                growth_rate = product.growth_rate_years[year - 1]
            else:
                growth_rate = product.terminal_growth_rate

            # 收入增长
            revenue *= (1 + growth_rate)

            # 营业利润
            operating_profit = revenue * product.operating_margin

            # 税后营业利润（NOPAT）
            nopat = operating_profit * (1 - tax_rate)

            # 折旧摊销
            depreciation = revenue * product.depreciation_ratio

            # 资本支出
            capex = revenue * product.capex_ratio

            # 营运资金变化
            wc_change = revenue * product.wc_change_ratio

            # 自由现金流 = NOPAT + 折旧 - 资本支出 - 营运资金变化
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
                'growth_rate': growth_rate,
            })

        return forecasts

    @staticmethod
    def calculate_product_valuation(
        product: ProductSegment,
        wacc: float,
        projection_years: int = 5,
        tax_rate: float = 0.25,
        terminal_method: str = "perpetuity"
    ) -> ProductValuationResult:
        """
        计算单个产品的DCF估值

        Args:
            product: 产品对象
            wacc: 加权平均资本成本
            projection_years: 预测年数
            tax_rate: 所得税率
            terminal_method: 终值计算方法

        Returns:
            产品估值结果
        """
        # 预测自由现金流
        fcf_forecasts = MultiProductValuation.forecast_product_cash_flows(
            product, projection_years, tax_rate
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
        terminal_growth_rate = product.terminal_growth_rate

        if terminal_method == "perpetuity":
            # 永续增长法：终值 = 最终FCF × (1 + g) / (WACC - g)
            terminal_value = final_fcf * (1 + terminal_growth_rate) / (wacc - terminal_growth_rate)
        else:
            # 退出倍数法（简化：使用终值倍数）
            terminal_multiple = 10.0  # 可以根据行业调整
            terminal_value = final_fcf * terminal_multiple

        # 终值折现
        pv_terminal = terminal_value / ((1 + wacc) ** projection_years)

        # 企业价值 = 预测期现值 + 终值现值
        enterprise_value = pv_forecasts + pv_terminal

        # 计算终值收入
        terminal_revenue = fcf_forecasts[-1]['revenue']

        # 计算收入复合增长率（CAGR）
        if projection_years > 0:
            revenue_cagr = (terminal_revenue / product.current_revenue) ** (1 / projection_years) - 1
        else:
            revenue_cagr = 0

        # 不再使用加权价值，直接使用企业价值
        # weighted_value 保留用于显示收入贡献占比，但不用于总价值计算

        return ProductValuationResult(
            product_name=product.name,
            revenue_weight=product.revenue_weight,
            pv_forecasts=pv_forecasts,
            pv_terminal=pv_terminal,
            enterprise_value=enterprise_value,
            weighted_value=enterprise_value,  # 改为直接使用企业价值
            fcf_forecasts=fcf_forecasts,
            current_revenue=product.current_revenue,
            terminal_revenue=terminal_revenue,
            revenue_cagr=revenue_cagr,
        )

    @staticmethod
    def consolidate_cash_flows(
        product_results: List[ProductValuationResult],
        projection_years: int
    ) -> List[Dict[str, float]]:
        """
        合并所有产品的现金流

        Args:
            product_results: 产品估值结果列表
            projection_years: 预测年数

        Returns:
            合并后的现金流预测
        """
        consolidated = []

        for year in range(1, projection_years + 1):
            year_revenue = 0
            year_operating_profit = 0
            year_nopat = 0
            year_depreciation = 0
            year_capex = 0
            year_wc_change = 0
            year_fcf = 0

            # 叠加所有产品的现金流（不再使用权重，因为current_revenue已经是绝对值）
            for result in product_results:
                if year <= len(result.fcf_forecasts):
                    forecast = result.fcf_forecasts[year - 1]
                    # 直接累加，不使用权重
                    year_revenue += forecast['revenue']
                    year_operating_profit += forecast['operating_profit']
                    year_nopat += forecast['nopat']
                    year_depreciation += forecast['depreciation']
                    year_capex += forecast['capex']
                    year_wc_change += forecast['wc_change']
                    year_fcf += forecast['fcf']

            consolidated.append({
                'year': year,
                'revenue': year_revenue,
                'operating_profit': year_operating_profit,
                'nopat': year_nopat,
                'depreciation': year_depreciation,
                'capex': year_capex,
                'wc_change': year_wc_change,
                'fcf': year_fcf,
            })

        return consolidated

    @staticmethod
    def multi_product_dcf_valuation(
        products: List[ProductSegment],
        company_beta: float = 1.0,
        tax_rate: float = 0.25,
        risk_free_rate: float = 0.03,
        market_risk_premium: float = 0.07,
        cost_of_debt: float = 0.05,
        target_debt_ratio: float = 0.3,
        total_debt: float = 0,
        cash_and_equivalents: float = 0,
        projection_years: int = 5,
        terminal_method: str = "perpetuity"
    ) -> MultiProductValuationResult:
        """
        多产品DCF估值

        Args:
            products: 产品列表
            company_beta: 公司整体贝塔系数（当产品无β时使用）
            tax_rate: 所得税率
            risk_free_rate: 无风险利率
            market_risk_premium: 市场风险溢价
            cost_of_debt: 债务成本
            target_debt_ratio: 目标资产负债率
            total_debt: 总债务（万元）
            cash_and_equivalents: 货币资金（万元）
            projection_years: 预测年数
            terminal_method: 终值计算方法

        Returns:
            多产品估值结果
        """
        # 验证产品权重总和
        total_weight = sum(p.revenue_weight for p in products)
        if abs(total_weight - 1.0) > 0.01:  # 允许1%的误差
            raise ValueError(f"产品权重总和应为1.0，当前为{total_weight:.2f}")

        # 计算公司整体WACC
        wacc = MultiProductValuation.calculate_wacc(
            company_beta, risk_free_rate, market_risk_premium,
            cost_of_debt, target_debt_ratio, tax_rate
        )

        # 对每个产品进行估值
        product_results = []
        total_revenue = 0
        revenue_by_product = {}

        for product in products:
            # 使用产品特定β或公司整体β
            product_beta = product.beta if product.beta is not None else company_beta

            # 计算产品WACC（可选：也可以使用公司整体WACC）
            product_wacc = MultiProductValuation.calculate_wacc(
                product_beta, risk_free_rate, market_risk_premium,
                cost_of_debt, target_debt_ratio, tax_rate
            )

            # 计算产品估值
            result = MultiProductValuation.calculate_product_valuation(
                product, wacc, projection_years, tax_rate, terminal_method
            )
            product_results.append(result)

            # 累加收入
            total_revenue += product.current_revenue
            revenue_by_product[product.name] = product.current_revenue

        # 叠加所有产品的价值（直接使用企业价值，不使用加权价值）
        total_enterprise_value = sum(r.enterprise_value for r in product_results)

        # 计算股权价值
        net_debt = total_debt - cash_and_equivalents
        total_equity_value = total_enterprise_value - net_debt

        # 价值分解（使用企业价值）
        value_breakdown = {r.product_name: r.enterprise_value for r in product_results}

        # 产品价值贡献分析（基于实际企业价值）
        product_contribution = []
        for result in product_results:
            contribution = result.enterprise_value / total_enterprise_value if total_enterprise_value > 0 else 0
            product_contribution.append({
                'product': result.product_name,
                'contribution': contribution,
                'contribution_pct': contribution * 100,
            })

        # 按贡献度排序
        product_contribution.sort(key=lambda x: x['contribution'], reverse=True)

        # 合并现金流
        consolidated_fcf = MultiProductValuation.consolidate_cash_flows(
            product_results, projection_years
        )

        return MultiProductValuationResult(
            total_enterprise_value=total_enterprise_value,
            total_equity_value=total_equity_value,
            wacc=wacc,
            product_results=product_results,
            value_breakdown=value_breakdown,
            total_revenue=total_revenue,
            revenue_by_product=revenue_by_product,
            product_contribution=product_contribution,
            consolidated_fcf_forecasts=consolidated_fcf,
        )


# ===== 辅助函数 =====

def validate_products(products: List[ProductSegment]) -> Tuple[bool, Optional[str]]:
    """
    验证产品列表

    Args:
        products: 产品列表

    Returns:
        (是否有效, 错误消息)
    """
    if not products:
        return False, "产品列表不能为空"

    if len(products) > 10:
        return False, "产品数量不能超过10个"

    # 检查权重总和
    total_weight = sum(p.revenue_weight for p in products)
    if abs(total_weight - 1.0) > 0.01:
        return False, f"产品权重总和应为100%，当前为{total_weight*100:.1f}%"

    # 检查每个产品的参数
    for i, product in enumerate(products):
        if not product.name:
            return False, f"产品{i+1}的名称不能为空"

        if product.current_revenue <= 0:
            return False, f"产品{product.name}的当前收入必须大于0"

        if product.revenue_weight <= 0 or product.revenue_weight > 1:
            return False, f"产品{product.name}的权重必须在0-1之间"

        if product.gross_margin < 0 or product.gross_margin > 1:
            return False, f"产品{product.name}的毛利率必须在0-1之间"

        if product.operating_margin < 0 or product.operating_margin > 1:
            return False, f"产品{product.name}的营业利润率必须在0-1之间"

        # 检查增长率列表
        if not product.growth_rate_years:
            return False, f"产品{product.name}的增长率列表不能为空"

        for j, growth in enumerate(product.growth_rate_years):
            if growth < -0.5 or growth > 1.0:
                return False, f"产品{product.name}的第{j+1}年增长率应在-50%到100%之间"

    return True, None


if __name__ == "__main__":
    # 测试多产品估值
    from models import ProductSegment

    # 创建3个产品
    products = [
        ProductSegment(
            name="云服务",
            description="云计算平台服务",
            current_revenue=50000,
            revenue_weight=0.6,
            growth_rate_years=[0.25, 0.20, 0.15, 0.10, 0.08],
            terminal_growth_rate=0.03,
            gross_margin=0.65,
            operating_margin=0.25,
            capex_ratio=0.08,
            wc_change_ratio=0.02,
            depreciation_ratio=0.04,
            beta=1.2,
        ),
        ProductSegment(
            name="软件许可",
            description="传统软件许可业务",
            current_revenue=30000,
            revenue_weight=0.3,
            growth_rate_years=[0.05, 0.03, 0.02, 0.02, 0.02],
            terminal_growth_rate=0.015,
            gross_margin=0.85,
            operating_margin=0.35,
            capex_ratio=0.02,
            wc_change_ratio=0.01,
            depreciation_ratio=0.01,
            beta=0.9,
        ),
        ProductSegment(
            name="技术服务",
            description="技术咨询与实施服务",
            current_revenue=20000,
            revenue_weight=0.1,
            growth_rate_years=[0.10, 0.08, 0.06, 0.05, 0.05],
            terminal_growth_rate=0.02,
            gross_margin=0.40,
            operating_margin=0.15,
            capex_ratio=0.03,
            wc_change_ratio=0.05,
            depreciation_ratio=0.02,
            beta=1.0,
        ),
    ]

    # 验证产品
    is_valid, error_msg = validate_products(products)
    if not is_valid:
        print(f"产品验证失败: {error_msg}")
    else:
        print("产品验证通过")

        # 进行估值
        result = MultiProductValuation.multi_product_dcf_valuation(
            products=products,
            company_beta=1.1,
            tax_rate=0.15,
            risk_free_rate=0.03,
            market_risk_premium=0.07,
            cost_of_debt=0.045,
            target_debt_ratio=0.2,
            total_debt=10000,
            cash_and_equivalents=5000,
            projection_years=5,
        )

        print(f"\n=== 多产品DCF估值结果 ===")
        print(f"企业价值: {result.total_enterprise_value/10000:.2f} 亿元")
        print(f"股权价值: {result.total_equity_value/10000:.2f} 亿元")
        print(f"WACC: {result.wacc:.2%}")

        print(f"\n=== 分产品估值明细 ===")
        for pr in result.product_results:
            print(f"\n{pr.product_name}:")
            print(f"  收入占比: {pr.revenue_weight:.1%}")
            print(f"  企业价值: {pr.enterprise_value/10000:.2f} 亿元")
            print(f"  加权价值: {pr.weighted_value/10000:.2f} 亿元")
            print(f"  收入CAGR: {pr.revenue_cagr:.2%}")
            print(f"  当前收入: {pr.current_revenue/10000:.2f} 亿元")
            print(f"  终值收入: {pr.terminal_revenue/10000:.2f} 亿元")

        print(f"\n=== 价值贡献 ===")
        for pc in result.product_contribution:
            print(f"{pc['product']}: {pc['contribution_pct']:.1f}%")
