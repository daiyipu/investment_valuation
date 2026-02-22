"""
其他估值方法模块
包含风险投资法（VC法）、成本法/净资产法、交易对价参考法等
"""
from typing import Optional, Dict, Any, List
from core.models import Company, ValuationResult


class OtherValuationMethods:
    """其他估值方法类"""

    @staticmethod
    def vc_method(
        company: Company,
        exit_valuation: float,
        target_return_multiple: float = 10.0,
        investment_years: int = 5,
        exit_method: str = "PE",
        exit_multiple: Optional[float] = None
    ) -> ValuationResult:
        """
        风险投资法估值（倒推法）

        从预期退出时的估值倒推当前投后估值

        Args:
            company: 目标公司
            exit_valuation: 预期退出时的估值（万元）
            target_return_multiple: 目标回报倍数
            investment_years: 投资年限
            exit_method: 退出估值方法
            exit_multiple: 退出倍数（如提供则用于计算退出估值）

        Returns:
            估值结果
        """
        # 如果提供退出倍数，重新计算退出估值
        if exit_multiple:
            if exit_method == "PE" and company.net_income > 0:
                # 假设退出时的净利润按当前增长率增长
                future_net_income = company.net_income * ((1 + company.growth_rate) ** investment_years)
                exit_valuation = future_net_income * exit_multiple
            elif exit_method == "PS" and company.revenue > 0:
                future_revenue = company.revenue * ((1 + company.growth_rate) ** investment_years)
                exit_valuation = future_revenue * exit_multiple

        # 倒推当前估值
        current_valuation = exit_valuation / target_return_multiple

        return ValuationResult(
            method="VC法",
            value=current_valuation,
            details={
                'exit_valuation': exit_valuation,
                'target_return_multiple': target_return_multiple,
                'investment_years': investment_years,
                'exit_method': exit_method,
                'exit_multiple': exit_multiple,
                'implied_irr': (target_return_multiple ** (1 / investment_years)) - 1,
            },
            assumptions={
                'growth_rate': company.growth_rate,
            }
        )

    @staticmethod
    def vc_method_with_future_projection(
        company: Company,
        projection_years: int = 5,
        target_pe: float = 20.0,
        target_return_multiple: float = 10.0,
        margin_improvement: float = 0.0
    ) -> ValuationResult:
        """
        带预测的风险投资法

        基于未来盈利预测和目标退出倍数进行估值

        Args:
            company: 目标公司
            projection_years: 预测年数
            target_pe: 目标退出时P/E倍数
            target_return_multiple: 目标回报倍数
            margin_improvement: 利润率年提升幅度

        Returns:
            估值结果
        """
        # 预测退出时净利润
        future_net_income = company.net_income

        for year in range(projection_years):
            # 收入增长
            future_net_income *= (1 + company.growth_rate)
            # 利润率提升
            if margin_improvement > 0:
                future_net_income *= (1 + margin_improvement)

        # 计算退出估值
        exit_valuation = future_net_income * target_pe

        # 倒推当前估值
        current_valuation = exit_valuation / target_return_multiple

        return ValuationResult(
            method="VC法（预测）",
            value=current_valuation,
            details={
                'future_net_income': future_net_income,
                'target_pe': target_pe,
                'exit_valuation': exit_valuation,
                'target_return_multiple': target_return_multiple,
                'projection_years': projection_years,
            },
            assumptions={
                'growth_rate': company.growth_rate,
                'margin_improvement': margin_improvement,
            }
        )

    @staticmethod
    def cost_method(
        company: Company,
        intangible_asset_value: float = 0,
        goodwill_value: float = 0,
        adjustment_factor: float = 1.0
    ) -> ValuationResult:
        """
        成本法/净资产法估值

        基于公司净资产价值进行估值，适用于资产密集型企业

        Args:
            company: 目标公司
            intangible_asset_value: 无形资产评估价值
            goodwill_value: 商誉价值
            adjustment_factor: 调整系数

        Returns:
            估值结果
        """
        if not company.net_assets:
            raise ValueError("缺少净资产数据")

        # 基础净资产价值
        base_value = company.net_assets

        # 加上无形资产和商誉
        adjusted_net_assets = base_value + intangible_asset_value + goodwill_value

        # 应用调整系数
        final_value = adjusted_net_assets * adjustment_factor

        return ValuationResult(
            method="成本法/净资产法",
            value=final_value,
            details={
                'net_assets': company.net_assets,
                'intangible_asset_value': intangible_asset_value,
                'goodwill_value': goodwill_value,
                'adjusted_net_assets': adjusted_net_assets,
                'adjustment_factor': adjustment_factor,
                'price_to_book_ratio': final_value / company.net_assets if company.net_assets > 0 else None,
            },
            assumptions={}
        )

    @staticmethod
    def adjusted_net_asset_method(
        company: Company,
        asset_adjustments: Dict[str, float],
        liability_adjustments: Dict[str, float]
    ) -> ValuationResult:
        """
        调整净资产法

        对各项资产和负债进行公允价值调整

        Args:
            company: 目标公司
            asset_adjustments: 资产调整字典 {资产名称: 调整金额}
            liability_adjustments: 负债调整字典

        Returns:
            估值结果
        """
        # 计算调整后的净资产
        total_asset_adjustment = sum(asset_adjustments.values())
        total_liability_adjustment = sum(liability_adjustments.values())

        if company.net_assets:
            adjusted_net_assets = company.net_assets + total_asset_adjustment - total_liability_adjustment
        else:
            raise ValueError("缺少净资产数据")

        return ValuationResult(
            method="调整净资产法",
            value=adjusted_net_assets,
            details={
                'original_net_assets': company.net_assets,
                'asset_adjustments': asset_adjustments,
                'liability_adjustments': liability_adjustments,
                'total_asset_adjustment': total_asset_adjustment,
                'total_liability_adjustment': total_liability_adjustment,
                'adjusted_net_assets': adjusted_net_assets,
            },
            assumptions={}
        )

    @staticmethod
    def transaction_comparable(
        company: Company,
        transactions: List[Dict[str, Any]]
    ) -> ValuationResult:
        """
        交易对价参考法

        参考近期类似交易的估值倍数

        Args:
            company: 目标公司
            transactions: 交易列表，每个交易包含：
                - company_name: 交易公司名称
                - deal_value: 交易金额
                - metric_value: 财务指标值（如收入、利润）
                - multiple: 交易倍数
                - deal_date: 交易日期
                - stage: 交易阶段

        Returns:
            估值结果
        """
        if not transactions:
            raise ValueError("缺少交易数据")

        # 提取交易倍数
        multiples = [t['multiple'] for t in transactions if t.get('multiple')]
        if not multiples:
            raise ValueError("交易数据中缺少倍数信息")

        import numpy as np

        # 计算平均/中位数倍数
        avg_multiple = np.mean(multiples)
        median_multiple = np.median(multiples)

        # 应用到目标公司
        # 优先使用净利润，其次收入
        if company.net_income > 0:
            metric_value = company.net_income
            metric_name = "净利润"
        elif company.revenue > 0:
            metric_value = company.revenue
            metric_name = "营业收入"
        else:
            raise ValueError("缺少可用于估值的财务指标")

        # 使用中位数倍数估值
        valuation = metric_value * median_multiple

        return ValuationResult(
            method="交易对价参考法",
            value=valuation,
            details={
                'transaction_count': len(transactions),
                'avg_multiple': avg_multiple,
                'median_multiple': median_multiple,
                'multiple_range': (min(multiples), max(multiples)),
                'metric_used': metric_name,
                'metric_value': metric_value,
                'transactions': transactions,
            },
            assumptions={}
        )

    @staticmethod
    def first_chicago_method(
        company: Company,
        success_scenario: Dict[str, Any],
        failure_scenario: Dict[str, Any],
        probability_of_success: float = 0.3
    ) -> ValuationResult:
        """
        第一芝加哥法

        针对早期项目，考虑成功和失败两种情景的加权估值

        Args:
            company: 目标公司
            success_scenario: 成功情景估值
            failure_scenario: 失败情景估值
            probability_of_success: 成功概率

        Returns:
            估值结果
        """
        success_value = success_scenario.get('value', 0)
        failure_value = failure_scenario.get('value', 0)

        # 期望价值
        expected_value = (success_value * probability_of_success +
                         failure_value * (1 - probability_of_success))

        return ValuationResult(
            method="第一芝加哥法",
            value=expected_value,
            details={
                'success_value': success_value,
                'failure_value': failure_value,
                'probability_of_success': probability_of_success,
                'success_scenario': success_scenario,
                'failure_scenario': failure_scenario,
            },
            assumptions={}
        )

    @staticmethod
    def sum_of_parts_valuation(
        company: Company,
        business_units: List[Dict[str, Any]]
    ) -> ValuationResult:
        """
        分部加总法

        将公司各业务部门分别估值后加总

        Args:
            company: 目标公司
            business_units: 业务单元列表，每个包含：
                - name: 业务名称
                - revenue: 收入
                - multiple: 估值倍数
                - value: 估值（如有直接估值）

        Returns:
            估值结果
        """
        parts_value = 0
        parts_details = []

        for unit in business_units:
            if unit.get('value'):
                unit_value = unit['value']
            elif unit.get('revenue') and unit.get('multiple'):
                unit_value = unit['revenue'] * unit['multiple']
            else:
                continue

            parts_value += unit_value
            parts_details.append({
                'name': unit['name'],
                'value': unit_value,
                'revenue': unit.get('revenue'),
                'multiple': unit.get('multiple'),
            })

        # 减去公司层面成本（协同效应折扣）
        corporate_discount = parts_value * 0.1  # 假设10%的公司层面成本
        final_value = parts_value - corporate_discount

        return ValuationResult(
            method="分部加总法",
            value=final_value,
            details={
                'parts_value': parts_value,
                'corporate_discount': corporate_discount,
                'business_units': parts_details,
            },
            assumptions={}
        )


# ===== 辅助函数 =====

def analyze_stage_appropriate_valuation(
    company: Company
) -> str:
    """
    根据公司阶段推荐估值方法

    Args:
        company: 目标公司

    Returns:
        推荐的估值方法
    """
    stage = company.stage

    if stage.value == "早期":
        return "VC法、交易对价法"
    elif stage.value == "成长期":
        return "P/S法、DCF法、VC法"
    elif stage.value == "成熟期":
        return "P/E法、DCF法、EV/EBITDA法"
    else:  # 上市公司
        return "P/E法、P/B法、EV/EBITDA法、DCF法"


# ===== 使用示例 =====

if __name__ == "__main__":
    from models import Company, CompanyStage

    print("=" * 60)
    print("其他估值方法示例")
    print("=" * 60)

    # 早期项目
    early_company = Company(
        name="早期科技公司",
        industry="人工智能",
        stage=CompanyStage.EARLY,
        revenue=2000,  # 2000万
        net_income=-500,  # 亏损
        net_assets=3000,
        growth_rate=0.50,
    )

    print(f"\n【公司】{early_company.name}")
    print(f"推荐估值方法: {analyze_stage_appropriate_valuation(early_company)}")

    # VC法估值
    vc_result = OtherValuationMethods.vc_method_with_future_projection(
        early_company,
        projection_years=5,
        target_pe=25.0,
        target_return_multiple=15.0,
    )
    print(f"\n{vc_result}")
    print(f"  退出估值: {vc_result.details['exit_valuation']/10000:.2f}亿元")
    print(f"  当前估值: {vc_result.value/10000:.2f}亿元")

    # 成长期项目
    growth_company = Company(
        name="成长期公司",
        industry="软件服务",
        stage=CompanyStage.GROWTH,
        revenue=20000,  # 2亿
        net_income=3000,  # 3000万
        net_assets=10000,
        growth_rate=0.40,
    )

    # 交易对价法
    transactions = [
        {'company_name': 'A公司', 'deal_value': 120000, 'metric_value': 6000, 'multiple': 20, 'stage': 'B轮'},
        {'company_name': 'B公司', 'deal_value': 80000, 'metric_value': 4000, 'multiple': 20, 'stage': 'B轮'},
        {'company_name': 'C公司', 'deal_value': 150000, 'metric_value': 5000, 'multiple': 30, 'stage': 'C轮'},
    ]

    trans_result = OtherValuationMethods.transaction_comparable(growth_company, transactions)
    print(f"\n{trans_result}")
    print(f"  参考交易: {trans_result.details['transaction_count']}笔")
    print(f"  中位数倍数: {trans_result.details['median_multiple']:.1f}x")

    # 第一芝加哥法
    success_scenario = {'value': 200000}  # 20亿
    failure_scenario = {'value': 10000}   # 1亿（清算价值）

    chicago_result = OtherValuationMethods.first_chicago_method(
        early_company,
        success_scenario,
        failure_scenario,
        probability_of_success=0.3
    )
    print(f"\n{chicago_result}")
    print(f"  成功情景: {success_scenario['value']/10000:.2f}亿元")
    print(f"  失败情景: {failure_scenario['value']/10000:.2f}亿元")
    print(f"  成功概率: {probability_of_success:.0%}")
    print(f"  期望价值: {chicago_result.value/10000:.2f}亿元")
