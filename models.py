"""
投资估值系统 - 核心数据模型
"""
from dataclasses import dataclass, field
from typing import Optional, List, Dict, Any
from datetime import datetime
from enum import Enum


class CompanyStage(Enum):
    """公司发展阶段"""
    EARLY = "早期"      # 天使轮、A轮
    GROWTH = "成长期"   # B轮、C轮
    MATURE = "成熟期"   # D轮及以上、Pre-IPO
    PUBLIC = "上市公司"


@dataclass
class Company:
    """公司基本信息和财务数据"""
    # 基本信息
    name: str                          # 公司名称
    industry: str                      # 所属行业
    stage: CompanyStage                # 发展阶段

    # 财务数据（单位：万元）
    revenue: float                     # 营业收入（最近年度）
    net_income: float                  # 净利润（最近年度）
    ebitda: Optional[float] = None     # 息税折旧摊销前利润
    gross_profit: Optional[float] = None  # 毛利润
    operating_cash_flow: Optional[float] = None  # 经营现金流

    # 资产负债数据（单位：万元）
    total_assets: Optional[float] = None    # 总资产
    net_assets: Optional[float] = None      # 净资产（股东权益）
    total_debt: Optional[float] = None      # 总债务
    cash_and_equivalents: Optional[float] = None  # 货币资金

    # 衍生财务数据
    net_debt: Optional[float] = field(init=False)  # 净债务 = 总债务 - 现金

    # 预测参数
    growth_rate: float = 0.15           # 预期收入增长率
    margin: Optional[float] = None      # 毛利率
    operating_margin: Optional[float] = None  # 营业利润率
    tax_rate: float = 0.25              # 所得税率

    # DCF相关参数
    beta: float = 1.0                   # 贝塔系数
    risk_free_rate: float = 0.03        # 无风险利率
    market_risk_premium: float = 0.07   # 市场风险溢价
    cost_of_debt: float = 0.05          # 债务成本
    target_debt_ratio: float = 0.3      # 目标资产负债率

    # 终值参数
    terminal_growth_rate: float = 0.025 # 永续增长率

    # 其他
    year: int = 2024                    # 财务数据年份
    description: Optional[str] = None   # 公司描述
    employee_count: Optional[int] = None  # 员工数量

    def __post_init__(self):
        """计算衍生字段"""
        # 计算净债务
        if self.total_debt is not None and self.cash_and_equivalents is not None:
            self.net_debt = self.total_debt - self.cash_and_equivalents
        else:
            self.net_debt = 0

        # 如果没有EBITDA，尝试计算
        if self.ebitda is None and self.operating_cash_flow is not None:
            # 简化估算：EBITDA ≈ 经营现金流 + 税费 + 利息
            # 这里用营业利润近似
            if self.operating_margin:
                operating_profit = self.revenue * self.operating_margin
                self.ebitda = operating_profit * 1.2  # 粗略估算

    @property
    def pe_ratio_implied(self) -> Optional[float]:
        """隐含市盈率（如果已知估值）"""
        if hasattr(self, '_market_cap') and self._market_cap and self.net_income > 0:
            return self._market_cap / self.net_income
        return None

    @property
    def ps_ratio_implied(self) -> Optional[float]:
        """隐含市销率"""
        if hasattr(self, '_market_cap') and self._market_cap and self.revenue > 0:
            return self._market_cap / self.revenue
        return None

    @property
    def pb_ratio_implied(self) -> Optional[float]:
        """隐含市净率"""
        if hasattr(self, '_market_cap') and self._market_cap and self.net_assets and self.net_assets > 0:
            return self._market_cap / self.net_assets
        return None

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        result = {}
        for key, value in self.__dict__.items():
            if isinstance(value, CompanyStage):
                result[key] = value.value
            elif isinstance(value, datetime):
                result[key] = value.isoformat()
            else:
                result[key] = value
        return result


@dataclass
class Comparable:
    """可比公司数据"""
    name: str                    # 公司名称
    ts_code: Optional[str] = None    # Tushare代码
    industry: str = ""           # 所属行业

    # 市场数据（单位：万元）
    market_cap: Optional[float] = None  # 总市值

    # 财务数据
    revenue: float = 0           # 营业收入
    net_income: float = 0        # 净利润
    net_assets: float = 0        # 净资产
    ebitda: Optional[float] = None

    # 估值倍数
    pe_ratio: Optional[float] = None   # 市盈率（TTM）
    ps_ratio: Optional[float] = None   # 市销率
    pb_ratio: Optional[float] = None   # 市净率
    ev_ebitda: Optional[float] = None  # EV/EBITDA

    # 其他指标
    growth_rate: Optional[float] = None  # 收入增长率
    year: int = 2024

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'name': self.name,
            'ts_code': self.ts_code,
            'industry': self.industry,
            'market_cap': self.market_cap,
            'revenue': self.revenue,
            'net_income': self.net_income,
            'net_assets': self.net_assets,
            'ebitda': self.ebitda,
            'pe_ratio': self.pe_ratio,
            'ps_ratio': self.ps_ratio,
            'pb_ratio': self.pb_ratio,
            'ev_ebitda': self.ev_ebitda,
            'growth_rate': self.growth_rate,
            'year': self.year,
        }


@dataclass
class ValuationResult:
    """估值结果容器"""
    method: str                    # 估值方法
    value: float                   # 估值结果（股权价值，单位：万元）
    currency: str = "CNY"          # 货币单位

    # 估值区间
    value_low: Optional[float] = None   # 估值下限
    value_high: Optional[float] = None  # 估值上限

    # 详细信息
    details: Dict[str, Any] = field(default_factory=dict)

    # 元数据
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())
    assumptions: Dict[str, Any] = field(default_factory=dict)

    @property
    def value_mid(self) -> float:
        """估值中值"""
        if self.value_low and self.value_high:
            return (self.value_low + self.value_high) / 2
        return self.value

    @property
    def range_width_pct(self) -> Optional[float]:
        """估值区间宽度百分比"""
        if self.value_low and self.value_high and self.value_mid > 0:
            return (self.value_high - self.value_low) / self.value_mid
        return None

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'method': self.method,
            'value': self.value,
            'value_low': self.value_low,
            'value_high': self.value_high,
            'value_mid': self.value_mid,
            'range_width_pct': self.range_width_pct,
            'currency': self.currency,
            'details': self.details,
            'timestamp': self.timestamp,
            'assumptions': self.assumptions,
        }

    def __str__(self) -> str:
        """字符串表示"""
        if self.value_low and self.value_high:
            return f"{self.method}: {self.value_low/10000:.2f} - {self.value_high/10000:.2f} 亿元"
        return f"{self.method}: {self.value/10000:.2f} 亿元"


@dataclass
class ScenarioConfig:
    """情景配置"""
    name: str                            # 情景名称
    revenue_growth_adj: float = 0.0      # 收入增长率调整（如+0.1表示增加10%）
    margin_adj: float = 0.0              # 毛利率调整
    wacc_adj: float = 0.0                # WACC调整
    terminal_growth_adj: float = 0.0     # 终值增长率调整
    cost_reduction_adj: float = 0.0      # 成本降低比例

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'name': self.name,
            'revenue_growth_adj': self.revenue_growth_adj,
            'margin_adj': self.margin_adj,
            'wacc_adj': self.wacc_adj,
            'terminal_growth_adj': self.terminal_growth_adj,
            'cost_reduction_adj': self.cost_reduction_adj,
        }


# 预定义情景
SCENARIOS = {
    'base': ScenarioConfig(name='基准情景'),
    'bull': ScenarioConfig(
        name='乐观情景',
        revenue_growth_adj=0.2,      # 收入增长+20%
        margin_adj=0.05,             # 毛利率+5%
        wacc_adj=-0.01,              # WACC-1%
        terminal_growth_adj=0.005,   # 终值增长+0.5%
    ),
    'bear': ScenarioConfig(
        name='悲观情景',
        revenue_growth_adj=-0.2,     # 收入增长-20%
        margin_adj=-0.05,            # 毛利率-5%
        wacc_adj=0.02,               # WACC+2%
        terminal_growth_adj=-0.005,  # 终值增长-0.5%
    ),
}


@dataclass
class StressTestResult:
    """压力测试结果"""
    test_name: str                      # 测试名称
    scenario_description: str            # 情景描述
    base_value: float                    # 基准估值
    stressed_value: float                # 压力情景估值
    change_pct: float                    # 变化百分比

    # 详细信息
    details: Dict[str, Any] = field(default_factory=dict)

    @property
    def downside_protection(self) -> float:
        """下行保护（损失百分比）"""
        return min(0, self.change_pct)

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'test_name': self.test_name,
            'scenario_description': self.scenario_description,
            'base_value': self.base_value,
            'stressed_value': self.stressed_value,
            'change_pct': self.change_pct,
            'downside_protection': self.downside_protection,
            'details': self.details,
        }


@dataclass
class MonteCarloResult:
    """蒙特卡洛模拟结果"""
    iterations: int                      # 迭代次数
    values: List[float]                  # 所有模拟值

    # 统计结果
    mean: Optional[float] = field(init=False)
    median: Optional[float] = field(init=False)
    std: Optional[float] = field(init=False)
    min_value: Optional[float] = field(init=False)
    max_value: Optional[float] = field(init=False)
    percentiles: Dict[str, float] = field(default_factory=dict)

    def __post_init__(self):
        """计算统计量"""
        import numpy as np
        arr = np.array(self.values)
        self.mean = float(np.mean(arr))
        self.median = float(np.median(arr))
        self.std = float(np.std(arr))
        self.min_value = float(np.min(arr))
        self.max_value = float(np.max(arr))
        self.percentiles = {
            'p5': float(np.percentile(arr, 5)),
            'p10': float(np.percentile(arr, 10)),
            'p25': float(np.percentile(arr, 25)),
            'p75': float(np.percentile(arr, 75)),
            'p90': float(np.percentile(arr, 90)),
            'p95': float(np.percentile(arr, 95)),
        }

    @property
    def confidence_interval_90(self) -> tuple:
        """90%置信区间"""
        return (self.percentiles['p5'], self.percentiles['p95'])

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        import numpy as np

        # 生成直方图分布数据（用于前端图表展示）
        arr = np.array(self.values)
        hist, bin_edges = np.histogram(arr, bins=30)
        distribution = []
        for i in range(len(hist)):
            distribution.append({
                'bin_lower': float(bin_edges[i]),
                'bin_upper': float(bin_edges[i + 1]),
                'count': int(hist[i])
            })

        return {
            'iterations': self.iterations,
            'mean': self.mean,
            'median': self.median,
            'std': self.std,
            'min_value': self.min_value,
            'max_value': self.max_value,
            'percentiles': self.percentiles,
            'confidence_interval_90': self.confidence_interval_90,
            'distribution': distribution,
            'percentile_5': self.percentiles['p5'],
            'percentile_95': self.percentiles['p95'],
        }


@dataclass
class ProductSegment:
    """产品/业务线数据模型"""
    # 基本信息
    name: str                                # 产品/业务线名称

    # 财务数据（当前年度）
    current_revenue: float                   # 当前收入（万元）
    revenue_weight: float                    # 收入占比（0-1）

    # 增长参数
    growth_rate_years: List[float] = field(default_factory=lambda: [0.15] * 5)  # 未来5年每年的增长率
    terminal_growth_rate: float = 0.025      # 永续增长率

    # 盈利能力参数
    gross_margin: float = 0.5                # 毛利率（0-1）
    operating_margin: float = 0.2            # 营业利润率（0-1）

    # 资本效率参数
    capex_ratio: float = 0.05                # 资本支出/收入
    wc_change_ratio: float = 0.02            # 营运资金变化/收入
    depreciation_ratio: float = 0.03         # 折旧摊销/收入

    # 其他
    description: Optional[str] = None        # 产品描述
    beta: Optional[float] = None             # 贝塔系数（如为空则使用公司整体β）

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'name': self.name,
            'description': self.description,
            'current_revenue': self.current_revenue,
            'revenue_weight': self.revenue_weight,
            'growth_rate_years': self.growth_rate_years,
            'terminal_growth_rate': self.terminal_growth_rate,
            'gross_margin': self.gross_margin,
            'operating_margin': self.operating_margin,
            'capex_ratio': self.capex_ratio,
            'wc_change_ratio': self.wc_change_ratio,
            'depreciation_ratio': self.depreciation_ratio,
            'beta': self.beta,
        }


@dataclass
class ProductValuationResult:
    """单个产品的估值结果"""
    product_name: str
    revenue_weight: float

    # 现值明细
    pv_forecasts: float  # 预测期现值
    pv_terminal: float   # 终值现值
    enterprise_value: float  # 企业价值

    # 价值贡献（等于企业价值，用于向后兼容）
    weighted_value: float  # 现在等于 enterprise_value（不再使用revenue_weight加权）

    # 现金流预测
    fcf_forecasts: List[Dict[str, float]]

    # 关键指标
    current_revenue: float
    terminal_revenue: float
    revenue_cagr: float  # 复合增长率

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'product_name': self.product_name,
            'revenue_weight': self.revenue_weight,
            'pv_forecasts': self.pv_forecasts,
            'pv_terminal': self.pv_terminal,
            'enterprise_value': self.enterprise_value,
            'weighted_value': self.weighted_value,
            'fcf_forecasts': self.fcf_forecasts,
            'current_revenue': self.current_revenue,
            'terminal_revenue': self.terminal_revenue,
            'revenue_cagr': self.revenue_cagr,
        }


@dataclass
class MultiProductValuationResult:
    """多产品估值总结果"""
    # 整体估值
    total_enterprise_value: float
    total_equity_value: float
    wacc: float

    # 分产品明细
    product_results: List[ProductValuationResult]

    # 价值分解
    value_breakdown: Dict[str, float]  # 产品名 -> 价值贡献

    # 汇总数据
    total_revenue: float
    revenue_by_product: Dict[str, float]

    # 分析指标
    product_contribution: List[Dict]  # 各产品价值贡献占比

    # 现金流汇总
    consolidated_fcf_forecasts: List[Dict]  # 叠加后的公司现金流

    # 元数据
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        return {
            'total_enterprise_value': self.total_enterprise_value,
            'total_equity_value': self.total_equity_value,
            'wacc': self.wacc,
            'product_results': [r.to_dict() for r in self.product_results],
            'value_breakdown': self.value_breakdown,
            'total_revenue': self.total_revenue,
            'revenue_by_product': self.revenue_by_product,
            'product_contribution': self.product_contribution,
            'consolidated_fcf_forecasts': self.consolidated_fcf_forecasts,
            'timestamp': self.timestamp,
        }


if __name__ == "__main__":
    # 测试数据模型
    company = Company(
        name="测试公司",
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
        margin=0.65,
    )
    print(f"公司: {company.name}")
    print(f"净债务: {company.net_debt} 万元")

    # 测试估值结果
    result = ValuationResult(
        method="P/E法",
        value=200000,  # 20亿元
        value_low=150000,
        value_high=250000,
    )
    print(f"\n{result}")
    print(f"估值区间宽度: {result.range_width_pct:.1%}")

    # 测试情景
    print(f"\n乐观情景: {SCENARIOS['bull'].name}")
