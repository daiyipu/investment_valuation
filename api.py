"""
投资估值系统 - FastAPI后端服务
提供RESTful API供前端调用
"""
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any
import os

from models import Company, Comparable, CompanyStage, ValuationResult, ScenarioConfig, ProductSegment
from relative_valuation import RelativeValuation
from absolute_valuation import AbsoluteValuation
from other_methods import OtherValuationMethods
from scenario_analysis import ScenarioAnalyzer
from stress_test import StressTester
from sensitivity_analysis import SensitivityAnalyzer
from multi_product_valuation import MultiProductValuation, validate_products
from database import DatabaseManager

# 初始化数据库
db = DatabaseManager()
db.create_tables()


# ===== 全局配置 =====

# Tushare Token存储（会话级）
TUSHARE_TOKEN = os.getenv("TUSHARE_TOKEN", "f2380d8761bcbf165f87b85f04ed105b1bdcf8721574562294671265")


# ===== FastAPI应用 =====

app = FastAPI(
    title="投资估值系统API",
    description="股权投资基金估值系统，提供相对估值、绝对估值、情景分析、压力测试等功能",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

# CORS配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",  # Vue开发服务器
        "http://localhost:3000",
        "http://localhost:8888",  # 静态文件服务器
        "http://127.0.0.1:5173",
        "http://127.0.0.1:3000",
        "http://127.0.0.1:8888",
        "*"  # 允许所有来源（开发环境）
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ===== Pydantic模型 =====

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


# ===== API端点 =====

@app.get("/")
async def root():
    """根路径"""
    return {
        "message": "投资估值系统API",
        "version": "1.0.0",
        "docs": "/docs",
        "history_api": "/api/history",  # 新增：历史记录API
    }


@app.get("/api/history", tags=["历史记录"])
async def get_history(
    limit: int = 50,
    offset: int = 0
):
    """
    获取估值历史记录

    Args:
        limit: 返回数量限制
        offset: 偏移量（分页）

    Returns:
        历史记录列表
    """
    history_list = db.get_history(limit=limit)

    return {
        "success": True,
        "total": len(history_list),
        "history": [h.dict() for h in history_list]
    }


@app.get("/api/history/{history_id}", tags=["历史记录"])
async def get_history_item(history_id: int):
    """
    获取单个历史记录详情

    Args:
        history_id: 历史记录ID

    Returns:
        历史记录详情（包含完整的结果数据）
    """
    import json
    from database import Session, ValuationHistory

    session = Session(db.get_engine())

    try:
        # 直接查询数据库模型以获取 notes 字段
        h = session.query(ValuationHistory).filter(ValuationHistory.id == history_id).first()

        if not h:
            raise HTTPException(status_code=404, detail="历史记录不存在")

        history_dict = {
            'id': h.id,
            'company_name': h.company_name,
            'industry': h.industry,
            'stage': h.stage,
            'revenue': h.revenue,
            'dcf_value': h.dcf_value,
            'dcf_wacc': h.dcf_wacc,
            'pe_value': h.pe_value,
            'ps_value': h.ps_value,
            'pb_value': h.pb_value,
            'ev_value': h.ev_value,
            'comparables_count': h.comparables_count,
            'created_at': h.created_at.isoformat() if h.created_at else None,
        }

        # 尝试从 notes 字段解析完整的 results 数据
        if h.notes:
            try:
                notes_data = json.loads(h.notes)
                if isinstance(notes_data, dict) and 'results' in notes_data:
                    history_dict['results'] = notes_data['results']
            except (json.JSONDecodeError, ValueError):
                # 兼容旧格式: "分析类型: {...json...}"
                try:
                    if ':' in h.notes:
                        _, json_str = h.notes.split(':', 1)
                        results_data = json.loads(json_str.strip())
                        history_dict['results'] = results_data
                except (json.JSONDecodeError, ValueError):
                    pass

        return {
            "success": True,
            "history": history_dict
        }
    finally:
        session.close()


@app.get("/health")
async def health_check():
    """健康检查"""
    return {"status": "healthy"}


# ===== 估值相关API =====

@app.post("/api/valuation/relative")
async def relative_valuation(request: RelativeValuationRequest):
    """
    相对估值法

    支持P/E、P/S、P/B、EV/EBITDA等多种相对估值方法
    """
    try:
        comp = request.company.to_company()
        comp_list = [c.to_comparable() for c in request.comparables]

        results = RelativeValuation.auto_comparable_analysis(comp, comp_list, request.methods)

        # 保存历史记录
        results_dict = {"results": {k: v.to_dict() for k, v in results.items()}}
        db.save_analysis_history('relative', request.company.dict(), results_dict)

        return {
            "success": True,
            "company": comp.name,
            "results": {k: v.to_dict() for k, v in results.items()},
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/valuation/absolute")
async def absolute_valuation(
    company: CompanyInput,
    projection_years: int = 5,
):
    """
    绝对估值法（DCF）

    使用现金流折现模型进行估值
    """
    try:
        comp = company.to_company()
        result = AbsoluteValuation.dcf_valuation(comp, projection_years=projection_years)

        # 保存历史记录
        result_dict = {"result": result.to_dict()}
        db.save_analysis_history('absolute', company.dict(), result_dict)

        return {
            "success": True,
            "company": comp.name,
            "result": result.to_dict(),
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/valuation/compare")
async def compare_valuation(
    company: CompanyInput,
    comparables: Optional[List[ComparableInput]] = None,
):
    """
    多种方法交叉验证估值

    综合相对估值和绝对估值结果
    """
    try:
        comp = company.to_company()

        results = {}

        # DCF估值
        dcf_result = AbsoluteValuation.dcf_valuation(comp)
        results["DCF"] = dcf_result.to_dict()

        # 如果有可比公司，进行相对估值
        if comparables:
            comp_list = [c.to_comparable() for c in comparables]
            relative_results = RelativeValuation.auto_comparable_analysis(comp, comp_list)
            for k, v in relative_results.items():
                results[k] = v.to_dict()

        return {
            "success": True,
            "company": comp.name,
            "results": results,
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# ===== 多产品估值API =====

@app.post("/api/valuation/multi-product-dcf")
async def multi_product_dcf(request: MultiProductValuationRequest):
    """
    多产品DCF估值

    支持对公司多个产品/业务线分别估值，然后叠加得到公司整体估值

    **使用场景**：
    - 公司有多个产品线，各产品增长率、利润率不同
    - 需要精细化估值，反映不同业务的差异
    - 需要分析各业务对整体价值的贡献

    **输入参数**：
    - company_name: 公司名称
    - industry: 所属行业
    - products: 产品列表（1-10个）
      - 每个产品包含：收入、占比、增长率、利润率等参数
    - 公司整体参数：税率、WACC参数、资本结构等

    **输出结果**：
    - 整体估值：企业价值、股权价值
    - 分产品估值明细：各产品的企业价值、加权价值
    - 价值贡献分析：各产品对整体价值的贡献占比
    - 分产品现金流预测
    - 合并后的公司现金流
    """
    try:
        # 转换产品列表
        products = [p.to_product_segment() for p in request.products]

        # 验证产品列表
        is_valid, error_msg = validate_products(products)
        if not is_valid:
            raise HTTPException(status_code=400, detail=error_msg)

        # 执行多产品估值
        result = MultiProductValuation.multi_product_dcf_valuation(
            products=products,
            company_beta=request.company_beta,
            tax_rate=request.tax_rate,
            risk_free_rate=request.risk_free_rate,
            market_risk_premium=request.market_risk_premium,
            cost_of_debt=request.cost_of_debt,
            target_debt_ratio=request.target_debt_ratio,
            total_debt=request.total_debt,
            cash_and_equivalents=request.cash_and_equivalents,
            projection_years=request.projection_years,
            terminal_method=request.terminal_method,
        )

        return {
            "success": True,
            "company": request.company_name,
            "industry": request.industry,
            "method": "多产品DCF法",
            "result": result.to_dict(),
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"估值失败: {str(e)}")


# ===== 情景分析API =====

class ScenarioAnalysisRequest(BaseModel):
    """情景分析请求模型"""
    company: CompanyInput
    scenarios: Optional[List[ScenarioInput]] = None


@app.post("/api/scenario/analyze")
async def scenario_analysis(request: ScenarioAnalysisRequest):
    """
    情景分析

    支持基准、乐观、悲观等多情景分析
    """
    try:
        print(f"DEBUG: Received scenario request")
        print(f"  company: {request.company.name}, growth_rate={request.company.growth_rate}")
        print(f"  scenarios: {request.scenarios}")

        comp = request.company.to_company()
        analyzer = ScenarioAnalyzer(comp)

        if request.scenarios is None:
            # 使用默认三情景
            results = analyzer.compare_scenarios()
        else:
            scenario_configs = [s.to_scenario_config() for s in request.scenarios]
            results = analyzer.compare_scenarios(scenarios=scenario_configs)

        # 转换结果格式
        formatted_results = {}
        for name, data in results.items():
            if name == 'statistics':
                formatted_results[name] = data
            else:
                formatted_results[name] = {
                    'scenario': data['scenario'].to_dict(),
                    'valuation': data['valuation'].to_dict(),
                    'value': data['value'],
                }

        # 保存历史记录
        db.save_analysis_history('scenario', request.company.dict(), formatted_results)

        return {
            "success": True,
            "company": comp.name,
            "results": formatted_results,
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# ===== 压力测试API =====

@app.post("/api/stress-test/revenue")
async def revenue_shock_test(
    company: CompanyInput,
    shocks: Optional[List[float]] = None,
):
    """
    收入冲击测试
    """
    try:
        comp = company.to_company()
        tester = StressTester(comp)

        if shocks is None:
            shocks = [-0.3, -0.2, -0.1]

        results = tester.revenue_shock_test(shocks)

        return {
            "success": True,
            "company": comp.name,
            "results": [r.to_dict() for r in results],
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/stress-test/monte-carlo")
async def monte_carlo_simulation(
    company: CompanyInput,
    iterations: int = Query(1000, ge=100, le=10000, description="迭代次数"),
):
    """
    蒙特卡洛模拟
    """
    try:
        comp = company.to_company()
        tester = StressTester(comp)

        result = tester.monte_carlo_simulation(iterations=iterations)

        return {
            "success": True,
            "company": comp.name,
            "result": result.to_dict(),
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/stress-test/full")
async def full_stress_test(company: CompanyInput):
    """
    综合压力测试
    """
    try:
        comp = company.to_company()
        tester = StressTester(comp)

        report = tester.generate_stress_report()

        return {
            "success": True,
            "company": comp.name,
            "report": report,
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# ===== 敏感性分析API =====

@app.post("/api/sensitivity/one-way")
async def one_way_sensitivity(
    company: CompanyInput,
    param_name: str,
    param_range: Optional[List[float]] = None,
):
    """
    单因素敏感性分析
    """
    try:
        comp = company.to_company()
        analyzer = SensitivityAnalyzer(comp)

        if param_range:
            range_tuple = (param_range[0], param_range[1])
        else:
            range_tuple = None

        result = analyzer.one_way_sensitivity(param_name, range_tuple)

        return {
            "success": True,
            "company": comp.name,
            "result": result,
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/sensitivity/tornado")
async def tornado_chart(
    company: CompanyInput,
    param_changes: Optional[Dict[str, float]] = None,
):
    """
    龙卷风图数据生成
    """
    try:
        comp = company.to_company()
        analyzer = SensitivityAnalyzer(comp)

        result = analyzer.tornado_chart_data(param_changes)

        return {
            "success": True,
            "company": comp.name,
            "result": result,
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/sensitivity/comprehensive")
async def comprehensive_sensitivity(company: CompanyInput):
    """
    综合敏感性分析
    """
    try:
        comp = company.to_company()
        analyzer = SensitivityAnalyzer(comp)

        results = analyzer.comprehensive_sensitivity_analysis()

        return {
            "success": True,
            "company": comp.name,
            "results": results,
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


# ===== 数据获取API =====

@app.post("/api/data/tushare/configure")
async def configure_tushare(token: str = Query(..., description="Tushare API Token")):
    """
    配置Tushare Token

    配置后可使用数据获取功能
    """
    global TUSHARE_TOKEN
    TUSHARE_TOKEN = token

    return {
        "success": True,
        "message": "Token已配置",
        "token_prefix": token[:8] + "***" if len(token) > 8 else "***"
    }


@app.get("/api/data/comparable/{industry}")
async def get_comparable_companies(
    industry: str,
    market_cap_min: Optional[float] = Query(None, description="最小市值（亿元）"),
    market_cap_max: Optional[float] = Query(None, description="最大市值（亿元）"),
    limit: int = Query(20, ge=1, le=100, description="返回数量")
):
    """
    获取可比公司列表

    从Tushare获取指定行业的上市公司数据
    """
    try:
        from data_fetcher import TushareDataFetcher

        if not TUSHARE_TOKEN:
            raise HTTPException(
                status_code=400,
                detail="未配置Tushare Token。请先通过 /api/data/tushare/configure 配置"
            )

        fetcher = TushareDataFetcher(TUSHARE_TOKEN)

        market_cap_range = None
        if market_cap_min and market_cap_max:
            market_cap_range = (market_cap_min, market_cap_max)

        companies = fetcher.get_comparable_companies(
            industry,
            market_cap_range=market_cap_range,
            limit=limit
        )

        return {
            "success": True,
            "industry": industry,
            "count": len(companies),
            "companies": [c.to_dict() for c in companies],
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"获取数据失败: {str(e)}")


@app.get("/api/data/industry-multiples/{industry}")
async def get_industry_multiples(
    industry: str,
    method: str = Query("median", description="统计方法：mean或median")
):
    """
    获取行业估值倍数

    返回指定行业的平均P/E、P/S、P/B等估值倍数
    """
    try:
        from data_fetcher import TushareDataFetcher

        if not TUSHARE_TOKEN:
            raise HTTPException(
                status_code=400,
                detail="未配置Tushare Token"
            )

        fetcher = TushareDataFetcher(TUSHARE_TOKEN)
        multiples = fetcher.get_industry_multiples(industry, method=method)

        if not multiples:
            raise HTTPException(status_code=404, detail=f"未找到{industry}行业的数据")

        return {
            "success": True,
            "industry": industry,
            "method": method,
            "multiples": multiples,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"获取数据失败: {str(e)}")


@app.get("/api/data/stock/{ts_code}")
async def get_stock_financial_data(ts_code: str):
    """
    获取单个上市公司的财务数据

    从Tushare获取指定股票代码的完整财务数据，用于估值输入
    """
    try:
        from data_fetcher import TushareDataFetcher

        if not TUSHARE_TOKEN:
            raise HTTPException(
                status_code=400,
                detail="未配置Tushare Token。请先通过 /api/data/tushare/configure 配置"
            )

        fetcher = TushareDataFetcher(TUSHARE_TOKEN)
        data = fetcher.get_financial_metrics(ts_code)

        if not data:
            raise HTTPException(status_code=404, detail=f"未找到股票代码{ts_code}的数据")

        return {
            "success": True,
            "ts_code": ts_code,
            "data": data,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"获取数据失败: {str(e)}")


@app.post("/api/data/search")
async def search_companies(
    keywords: List[str],
    limit: int = Query(10, ge=1, le=50)
):
    """
    按关键词搜索公司

    根据公司名称关键词搜索相关上市公司
    """
    try:
        from data_fetcher import TushareDataFetcher

        if not TUSHARE_TOKEN:
            raise HTTPException(
                status_code=400,
                detail="未配置Tushare Token"
            )

        fetcher = TushareDataFetcher(TUSHARE_TOKEN)
        companies = fetcher.search_by_keywords(keywords, limit)

        return {
            "success": True,
            "keywords": keywords,
            "count": len(companies),
            "companies": [c.to_dict() for c in companies],
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"搜索失败: {str(e)}")


# ===== 启动服务器 =====

if __name__ == "__main__":
    import uvicorn

    print("启动投资估值系统API服务器...")
    print("API文档地址: http://localhost:8000/docs")
    print("ReDoc文档地址: http://localhost:8000/redoc")

    uvicorn.run(app, host="0.0.0.0", port=8000)
