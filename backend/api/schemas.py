"""
API 请求和响应的 Pydantic 模型
"""
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
import sys
import os

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from core.models import Company, Comparable, CompanyStage, ValuationResult, ScenarioConfig, ProductSegment


class CompanyInput(BaseModel):
    """公司信息输入模型"""
    name: str = Field(..., description="公司名称")
    industry: str = Field(..., description="所属行业")
    stage: str = Field(..., description="发展阶段（早期/成长期/成熟期/上市公司）")

    # 财务数据（单位：万元）
    revenue: float = Field(..., description="营业收入（万元）")
    net_income: float = Field(..., description="净利润（万元）")
    ebitda: Optional[float] = Field(None, description="息税折旧摊销前利润（万元）")
    gross_profit: Optional[float] = Field(None, description="毛利润（万元）")
    operating_cash_flow: Optional[float] = Field(None, description="经营现金流（万元）")

    # 资产负债数据
    total_assets: Optional[float] = Field(None, description="总资产（万元）")
    net_assets: Optional[float] = Field(None, description="净资产（万元）")
    total_debt: Optional[float] = Field(None, description="总债务（万元）")
    cash_and_equivalents: Optional[float] = Field(None, description="货币资金（万元）")

    # 预测参数
    growth_rate: float = Field(0.15, description="预期收入增长率")
    margin: Optional[float] = Field(None, description="毛利率")
    operating_margin: Optional[float] = Field(None, description="营业利润率")
    tax_rate: float = Field(0.25, description="所得税率")

    # DCF参数
    beta: float = Field(1.0, description="贝塔系数")
    risk_free_rate: float = Field(0.03, description="无风险利率")
    market_risk_premium: float = Field(0.07, description="市场风险溢价")
    cost_of_debt: float = Field(0.05, description="债务成本")
    target_debt_ratio: float = Field(0.3, description="目标资产负债率")
    terminal_growth_rate: float = Field(0.025, description="永续增长率")

    def to_company(self) -> Company:
        """转换为Company对象"""
        stage_map = {
            "早期": CompanyStage.EARLY,
            "成长期": CompanyStage.GROWTH,
            "成熟期": CompanyStage.MATURE,
            "上市公司": CompanyStage.PUBLIC,
        }

        return Company(
            name=self.name,
            industry=self.industry,
            stage=stage_map.get(self.stage, CompanyStage.GROWTH),
            revenue=self.revenue,
            net_income=self.net_income,
            ebitda=self.ebitda,
            gross_profit=self.gross_profit,
            operating_cash_flow=self.operating_cash_flow,
            total_assets=self.total_assets,
            net_assets=self.net_assets,
            total_debt=self.total_debt,
            cash_and_equivalents=self.cash_and_equivalents,
            growth_rate=self.growth_rate,
            margin=self.margin,
            operating_margin=self.operating_margin,
            tax_rate=self.tax_rate,
            beta=self.beta,
            risk_free_rate=self.risk_free_rate,
            market_risk_premium=self.market_risk_premium,
            cost_of_debt=self.cost_of_debt,
            target_debt_ratio=self.target_debt_ratio,
            terminal_growth_rate=self.terminal_growth_rate,
        )


class ComparableInput(BaseModel):
    """可比公司输入模型"""
    name: str
    ts_code: Optional[str] = None
    industry: str
    revenue: float
    net_income: float
    net_assets: float
    ebitda: Optional[float] = None
    pe_ratio: Optional[float] = None
    ps_ratio: Optional[float] = None
    pb_ratio: Optional[float] = None
    ev_ebitda: Optional[float] = None
    growth_rate: Optional[float] = None

    def to_comparable(self) -> Comparable:
        """转换为Comparable对象"""
        return Comparable(
            name=self.name,
            ts_code=self.ts_code,
            industry=self.industry,
            revenue=self.revenue,
            net_income=self.net_income,
            net_assets=self.net_assets,
            ebitda=self.ebitda,
            pe_ratio=self.pe_ratio,
            ps_ratio=self.ps_ratio,
            pb_ratio=self.pb_ratio,
            ev_ebitda=self.ev_ebitda,
            growth_rate=self.growth_rate,
        )


class ScenarioInput(BaseModel):
    """情景配置输入模型"""
    name: str
    revenue_growth_adj: float = 0.0
    margin_adj: float = 0.0
    wacc_adj: float = 0.0
    terminal_growth_adj: float = 0.0

    def to_scenario_config(self) -> ScenarioConfig:
        """转换为ScenarioConfig对象"""
        return ScenarioConfig(
            name=self.name,
            revenue_growth_adj=self.revenue_growth_adj,
            margin_adj=self.margin_adj,
            wacc_adj=self.wacc_adj,
            terminal_growth_adj=self.terminal_growth_adj,
        )


class RelativeValuationRequest(BaseModel):
    """相对估值请求模型（包装类）"""
    company: CompanyInput
    comparables: List[ComparableInput]
    methods: Optional[List[str]] = None


class ProductSegmentInput(BaseModel):
    """产品/业务线输入模型"""
    name: str = Field(..., description="产品/业务线名称")
    description: Optional[str] = Field(None, description="产品描述")

    # 财务数据（当前年度）
    current_revenue: float = Field(..., description="当前收入（万元）", gt=0)
    revenue_weight: float = Field(..., description="收入占比（0-1）", ge=0, le=1)

    # 增长参数
    growth_rate_years: List[float] = Field(
        default=[0.15, 0.15, 0.15, 0.15, 0.15],
        description="未来5年每年的增长率"
    )
    terminal_growth_rate: float = Field(0.025, description="永续增长率")

    # 盈利能力参数
    gross_margin: float = Field(0.5, description="毛利率（0-1）", ge=0, le=1)
    operating_margin: float = Field(0.2, description="营业利润率（0-1）", ge=0, le=1)

    # 资本效率参数
    capex_ratio: float = Field(0.05, description="资本支出/收入", ge=0)
    wc_change_ratio: float = Field(0.02, description="营运资金变化/收入", ge=0)
    depreciation_ratio: float = Field(0.03, description="折旧摊销/收入", ge=0)

    # 风险参数
    beta: Optional[float] = Field(None, description="贝塔系数（如为空则使用公司整体β）", ge=0)

    def to_product_segment(self) -> ProductSegment:
        """转换为ProductSegment对象"""
        return ProductSegment(
            name=self.name,
            description=self.description,
            current_revenue=self.current_revenue,
            revenue_weight=self.revenue_weight,
            growth_rate_years=self.growth_rate_years,
            terminal_growth_rate=self.terminal_growth_rate,
            gross_margin=self.gross_margin,
            operating_margin=self.operating_margin,
            capex_ratio=self.capex_ratio,
            wc_change_ratio=self.wc_change_ratio,
            depreciation_ratio=self.depreciation_ratio,
            beta=self.beta,
        )


class MultiProductValuationRequest(BaseModel):
    """多产品估值请求模型"""
    company_name: str = Field(..., description="公司名称")
    industry: str = Field(..., description="所属行业")
    products: List[ProductSegmentInput] = Field(
        ...,
        description="产品/业务线列表",
        min_items=1,
        max_items=10
    )

    # 公司整体参数
    tax_rate: float = Field(0.25, description="所得税率", ge=0, le=1)
    risk_free_rate: float = Field(0.03, description="无风险利率")
    market_risk_premium: float = Field(0.07, description="市场风险溢价")
    cost_of_debt: float = Field(0.05, description="债务成本")

    # 资本结构
    target_debt_ratio: float = Field(0.3, description="目标资产负债率", ge=0, le=1)
    total_debt: float = Field(0, description="总债务（万元）", ge=0)
    cash_and_equivalents: float = Field(0, description="货币资金（万元）", ge=0)

    # DCF参数
    projection_years: int = Field(5, description="预测年数", ge=3, le=10)
    terminal_method: str = Field("perpetuity", description="终值计算方法")

    # 公司整体β（当产品无β时使用）
    company_beta: float = Field(1.0, description="公司整体贝塔系数")


class ScenarioAnalysisRequest(BaseModel):
    """情景分析请求模型"""
    company: CompanyInput
    scenarios: Optional[List[ScenarioInput]] = None
    # 多产品模式支持
    products: Optional[List[ProductSegmentInput]] = None
    company_beta: Optional[float] = Field(None, description="公司整体贝塔系数（多产品模式使用）")
    tax_rate: Optional[float] = Field(0.25, description="税率（多产品模式使用）")
    total_debt: Optional[float] = Field(None, description="总债务（多产品模式使用，万元）")
    cash_and_equivalents: Optional[float] = Field(0, description="货币资金（多产品模式使用，万元）")
