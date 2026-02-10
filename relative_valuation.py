"""
相对估值法模块
包含P/E、P/S、P/B、EV/EBITDA等相对估值方法
"""
import numpy as np
from typing import List, Optional, Dict, Any, Union
from models import Company, Comparable, ValuationResult


class RelativeValuation:
    """相对估值计算类"""

    @staticmethod
    def pe_ratio_valuation(
        company: Company,
        comparable_pe_ratios: List[float],
        use_future_earnings: bool = True,
        discount_for_illiquidity: float = 0.0,
        control_premium: float = 0.0
    ) -> ValuationResult:
        """
        市盈率法估值

        Args:
            company: 目标公司
            comparable_pe_ratios: 可比公司P/E倍数列表
            use_future_earnings: 是否使用未来一年预测利润
            discount_for_illiquidity: 流动性折价（如0.2表示折价20%）
            control_premium: 控制权溢价（如0.1表示溢价10%）

        Returns:
            估值结果
        """
        if company.net_income <= 0:
            raise ValueError("净利润为负，不适用于P/E法")

        if not comparable_pe_ratios:
            raise ValueError("缺少可比公司P/E数据")

        # 计算可比公司P/E统计量
        pe_mean = np.mean(comparable_pe_ratios)
        pe_median = np.median(comparable_pe_ratios)
        pe_std = np.std(comparable_pe_ratios)
        pe_min = np.min(comparable_pe_ratios)
        pe_max = np.max(comparable_pe_ratios)

        # 计算目标公司净利润
        if use_future_earnings:
            # 使用未来一年预测利润
            earnings = company.net_income * (1 + company.growth_rate)
        else:
            earnings = company.net_income

        # 计算估值
        base_value = earnings * pe_mean

        # 应用调整
        total_adjustment = 1 - discount_for_illiquidity + control_premium
        adjusted_value = base_value * total_adjustment

        # 计算估值区间
        value_low = earnings * pe_min * total_adjustment
        value_high = earnings * pe_max * total_adjustment

        return ValuationResult(
            method="P/E法",
            value=adjusted_value,
            value_low=value_low,
            value_high=value_high,
            details={
                'pe_mean': pe_mean,
                'pe_median': pe_median,
                'pe_std': pe_std,
                'pe_range': (pe_min, pe_max),
                'earnings_used': earnings,
                'is_forward': use_future_earnings,
                'discount_for_illiquidity': discount_for_illiquidity,
                'control_premium': control_premium,
            },
            assumptions={
                'comparable_count': len(comparable_pe_ratios),
            }
        )

    @staticmethod
    def ps_ratio_valuation(
        company: Company,
        comparable_ps_ratios: List[float],
        use_future_revenue: bool = True,
        discount_for_illiquidity: float = 0.0,
        control_premium: float = 0.0
    ) -> ValuationResult:
        """
        市销率法估值

        Args:
            company: 目标公司
            comparable_ps_ratios: 可比公司P/S倍数列表
            use_future_revenue: 是否使用未来一年预测收入
            discount_for_illiquidity: 流动性折价
            control_premium: 控制权溢价

        Returns:
            估值结果
        """
        if company.revenue <= 0:
            raise ValueError("营业收入为负，不适用于P/S法")

        if not comparable_ps_ratios:
            raise ValueError("缺少可比公司P/S数据")

        # 计算P/S统计量
        ps_mean = np.mean(comparable_ps_ratios)
        ps_median = np.median(comparable_ps_ratios)
        ps_std = np.std(comparable_ps_ratios)
        ps_min = np.min(comparable_ps_ratios)
        ps_max = np.max(comparable_ps_ratios)

        # 计算目标公司收入
        if use_future_revenue:
            revenue = company.revenue * (1 + company.growth_rate)
        else:
            revenue = company.revenue

        # 计算估值
        base_value = revenue * ps_mean

        # 应用调整
        total_adjustment = 1 - discount_for_illiquidity + control_premium
        adjusted_value = base_value * total_adjustment

        # 计算估值区间
        value_low = revenue * ps_min * total_adjustment
        value_high = revenue * ps_max * total_adjustment

        return ValuationResult(
            method="P/S法",
            value=adjusted_value,
            value_low=value_low,
            value_high=value_high,
            details={
                'ps_mean': ps_mean,
                'ps_median': ps_median,
                'ps_std': ps_std,
                'ps_range': (ps_min, ps_max),
                'revenue_used': revenue,
                'is_forward': use_future_revenue,
                'discount_for_illiquidity': discount_for_illiquidity,
                'control_premium': control_premium,
            },
            assumptions={
                'comparable_count': len(comparable_ps_ratios),
            }
        )

    @staticmethod
    def pb_ratio_valuation(
        company: Company,
        comparable_pb_ratios: List[float],
        discount_for_illiquidity: float = 0.0,
        control_premium: float = 0.0
    ) -> ValuationResult:
        """
        市净率法估值

        Args:
            company: 目标公司
            comparable_pb_ratios: 可比公司P/B倍数列表
            discount_for_illiquidity: 流动性折价
            control_premium: 控制权溢价

        Returns:
            估值结果
        """
        if not company.net_assets or company.net_assets <= 0:
            raise ValueError("净资产为负或不适用，不适用于P/B法")

        if not comparable_pb_ratios:
            raise ValueError("缺少可比公司P/B数据")

        # 计算P/B统计量
        pb_mean = np.mean(comparable_pb_ratios)
        pb_median = np.median(comparable_pb_ratios)
        pb_std = np.std(comparable_pb_ratios)
        pb_min = np.min(comparable_pb_ratios)
        pb_max = np.max(comparable_pb_ratios)

        # 计算估值
        base_value = company.net_assets * pb_mean

        # 应用调整
        total_adjustment = 1 - discount_for_illiquidity + control_premium
        adjusted_value = base_value * total_adjustment

        # 计算估值区间
        value_low = company.net_assets * pb_min * total_adjustment
        value_high = company.net_assets * pb_max * total_adjustment

        return ValuationResult(
            method="P/B法",
            value=adjusted_value,
            value_low=value_low,
            value_high=value_high,
            details={
                'pb_mean': pb_mean,
                'pb_median': pb_median,
                'pb_std': pb_std,
                'pb_range': (pb_min, pb_max),
                'net_assets': company.net_assets,
                'discount_for_illiquidity': discount_for_illiquidity,
                'control_premium': control_premium,
            },
            assumptions={
                'comparable_count': len(comparable_pb_ratios),
            }
        )

    @staticmethod
    def ev_ebitda_valuation(
        company: Company,
        comparable_ev_ebitda: List[float],
        discount_for_illiquidity: float = 0.0,
        control_premium: float = 0.0
    ) -> ValuationResult:
        """
        企业价值倍数法估值（EV/EBITDA）

        Args:
            company: 目标公司
            comparable_ev_ebitda: 可比公司EV/EBITDA倍数列表
            discount_for_illiquidity: 流动性折价
            control_premium: 控制权溢价

        Returns:
            估值结果
        """
        if not company.ebitda or company.ebitda <= 0:
            raise ValueError("EBITDA为负或不适用，不适用于EV/EBITDA法")

        if not comparable_ev_ebitda:
            raise ValueError("缺少可比公司EV/EBITDA数据")

        # 计算EV/EBITDA统计量
        ev_mean = np.mean(comparable_ev_ebitda)
        ev_median = np.median(comparable_ev_ebitda)
        ev_std = np.std(comparable_ev_ebitda)
        ev_min = np.min(comparable_ev_ebitda)
        ev_max = np.max(comparable_ev_ebitda)

        # 计算企业价值
        enterprise_value = company.ebitda * ev_mean

        # 应用调整
        total_adjustment = 1 - discount_for_illiquidity + control_premium
        adjusted_ev = enterprise_value * total_adjustment

        # 计算股权价值: 股权价值 = 企业价值 - 净债务
        equity_value = adjusted_ev - company.net_debt
        equity_low = company.ebitda * ev_min * total_adjustment - company.net_debt
        equity_high = company.ebitda * ev_max * total_adjustment - company.net_debt

        return ValuationResult(
            method="EV/EBITDA法",
            value=equity_value,
            value_low=equity_low,
            value_high=equity_high,
            details={
                'ev_ebitda_mean': ev_mean,
                'ev_ebitda_median': ev_median,
                'ev_ebitda_std': ev_std,
                'ev_ebitda_range': (ev_min, ev_max),
                'ebitda': company.ebitda,
                'enterprise_value': adjusted_ev,
                'net_debt': company.net_debt,
                'discount_for_illiquidity': discount_for_illiquidity,
                'control_premium': control_premium,
            },
            assumptions={
                'comparable_count': len(comparable_ev_ebitda),
            }
        )

    @staticmethod
    def auto_comparable_analysis(
        company: Company,
        comparables: List[Comparable],
        methods: Optional[List[str]] = None,
        weight_pe: float = 0.3,
        weight_ps: float = 0.3,
        weight_pb: float = 0.2,
        weight_ev: float = 0.2
    ) -> Dict[str, ValuationResult]:
        """
        自动使用可比公司进行多种方法估值

        Args:
            company: 目标公司
            comparables: 可比公司列表
            methods: 要使用的估值方法列表，默认全部
            weight_pe: P/E法权重
            weight_ps: P/S法权重
            weight_pb: P/B法权重
            weight_ev: EV/EBITDA法权重

        Returns:
            各方法估值结果字典
        """
        if methods is None:
            methods = ['PE', 'PS', 'PB', 'EV']

        results = {}

        # 提取可比公司数据
        pe_ratios = [c.pe_ratio for c in comparables if c.pe_ratio and c.pe_ratio > 0]
        ps_ratios = [c.ps_ratio for c in comparables if c.ps_ratio and c.ps_ratio > 0]
        pb_ratios = [c.pb_ratio for c in comparables if c.pb_ratio and c.pb_ratio > 0]
        ev_ratios = [c.ev_ebitda for c in comparables if c.ev_ebitda and c.ev_ebitda > 0]

        # P/E法
        if 'PE' in methods and pe_ratios and company.net_income > 0:
            try:
                results['PE'] = RelativeValuation.pe_ratio_valuation(
                    company, pe_ratios
                )
            except ValueError as e:
                print(f"P/E法估值失败: {e}")

        # P/S法
        if 'PS' in methods and ps_ratios and company.revenue > 0:
            try:
                results['PS'] = RelativeValuation.ps_ratio_valuation(
                    company, ps_ratios
                )
            except ValueError as e:
                print(f"P/S法估值失败: {e}")

        # P/B法
        if 'PB' in methods and pb_ratios and company.net_assets and company.net_assets > 0:
            try:
                results['PB'] = RelativeValuation.pb_ratio_valuation(
                    company, pb_ratios
                )
            except ValueError as e:
                print(f"P/B法估值失败: {e}")

        # EV/EBITDA法
        if 'EV' in methods and ev_ratios and company.ebitda and company.ebitda > 0:
            try:
                results['EV'] = RelativeValuation.ev_ebitda_valuation(
                    company, ev_ratios
                )
            except ValueError as e:
                print(f"EV/EBITDA法估值失败: {e}")

        # 综合估值（加权平均）
        if len(results) > 1:
            weighted_value = 0
            total_weight = 0

            if 'PE' in results:
                weighted_value += results['PE'].value * weight_pe
                total_weight += weight_pe
            if 'PS' in results:
                weighted_value += results['PS'].value * weight_ps
                total_weight += weight_ps
            if 'PB' in results:
                weighted_value += results['PB'].value * weight_pb
                total_weight += weight_pb
            if 'EV' in results:
                weighted_value += results['EV'].value * weight_ev
                total_weight += weight_ev

            if total_weight > 0:
                weighted_value /= total_weight

                # 计算综合区间
                all_lows = [r.value_low for r in results.values() if r.value_low]
                all_highs = [r.value_high for r in results.values() if r.value_high]

                results['综合'] = ValuationResult(
                    method="综合估值（加权平均）",
                    value=weighted_value,
                    value_low=min(all_lows) if all_lows else None,
                    value_high=max(all_highs) if all_highs else None,
                    details={
                        'weights': {
                            'PE': weight_pe,
                            'PS': weight_ps,
                            'PB': weight_pb,
                            'EV': weight_ev,
                        },
                        'methods_used': list(results.keys()),
                    }
                )

        return results


# ===== 辅助函数 =====

def find_comparable_multiples(
    comparables: List[Comparable]
) -> Dict[str, List[float]]:
    """
    从可比公司列表中提取估值倍数

    Args:
        comparables: 可比公司列表

    Returns:
        估值倍数字典
    """
    return {
        'pe_ratios': [c.pe_ratio for c in comparables if c.pe_ratio and c.pe_ratio > 0],
        'ps_ratios': [c.ps_ratio for c in comparables if c.ps_ratio and c.ps_ratio > 0],
        'pb_ratios': [c.pb_ratio for c in comparables if c.pb_ratio and c.pb_ratio > 0],
        'ev_ebitda': [c.ev_ebitda for c in comparables if c.ev_ebitda and c.ev_ebitda > 0],
    }


def analyze_comparable_statistics(comparables: List[Comparable]) -> Dict[str, Any]:
    """
    分析可比公司统计数据

    Args:
        comparables: 可比公司列表

    Returns:
        统计分析结果
    """
    multiples = find_comparable_multiples(comparables)

    stats = {}
    for key, values in multiples.items():
        if values:
            stats[key] = {
                'count': len(values),
                'mean': np.mean(values),
                'median': np.median(values),
                'std': np.std(values),
                'min': np.min(values),
                'max': np.max(values),
            }

    return stats


# ===== 使用示例 =====

if __name__ == "__main__":
    from models import Company, Comparable, CompanyStage

    # 创建目标公司
    target = Company(
        name="目标科技公司",
        industry="软件服务",
        stage=CompanyStage.GROWTH,
        revenue=50000,  # 5亿
        net_income=8000,  # 8000万
        ebitda=12000,
        net_assets=20000,
        total_debt=5000,
        cash_and_equivalents=2000,
        growth_rate=0.25,
    )

    # 创建可比公司
    comp1 = Comparable(
        name="可比公司A",
        ts_code="000001.SZ",
        revenue=80000,
        net_income=15000,
        net_assets=30000,
        pe_ratio=25.5,
        ps_ratio=6.2,
        pb_ratio=4.8,
        ev_ebitda=18.3,
    )

    comp2 = Comparable(
        name="可比公司B",
        revenue=60000,
        net_income=10000,
        net_assets=25000,
        pe_ratio=22.0,
        ps_ratio=5.5,
        pb_ratio=4.2,
        ev_ebitda=16.5,
    )

    comp3 = Comparable(
        name="可比公司C",
        revenue=100000,
        net_income=20000,
        net_assets=40000,
        pe_ratio=28.0,
        ps_ratio=7.0,
        pb_ratio=5.5,
        ev_ebitda=20.0,
    )

    comparables = [comp1, comp2, comp3]

    # 自动估值分析
    results = RelativeValuation.auto_comparable_analysis(target, comparables)

    print("=" * 50)
    print(f"{'相对估值分析结果':^48}")
    print("=" * 50)

    for method, result in results.items():
        print(f"\n{result}")
        if result.details:
            print(f"  详情: {result.details}")

    # 统计分析
    print("\n" + "=" * 50)
    print("可比公司统计:")
    stats = analyze_comparable_statistics(comparables)
    for metric, stat in stats.items():
        print(f"\n{metric}:")
        print(f"  数量: {stat['count']}")
        print(f"  均值: {stat['mean']:.2f}")
        print(f"  中位数: {stat['median']:.2f}")
        print(f"  范围: {stat['min']:.2f} - {stat['max']:.2f}")
