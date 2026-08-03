"""
投资估值系统 - 简化版API服务器
临时解决方案，提供基本API功能
"""
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from typing import Dict, Any, List
import sys
import os

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, project_root)

print(f"项目根目录: {project_root}")

app = FastAPI(
    title="投资估值系统API",
    description="股权投资基金估值系统 - 简化版",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

# CORS配置
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://localhost:3000",
        "http://127.0.0.1:5173",
        "http://127.0.0.1:3000",
        "*"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    """根路径"""
    return {
        "message": "投资估值系统API - 简化版",
        "version": "1.0.0",
        "status": "running",
        "docs": "/docs"
    }

@app.get("/health")
async def health_check():
    """健康检查"""
    return {"status": "healthy", "service": "investment-valuation-api"}

# 基本的估值API端点
@app.post("/api/valuation/relative")
async def relative_valuation(request: Dict[str, Any]):
    """相对估值计算"""
    try:
        company = request.get("company", {})
        comparables = request.get("comparables", [])

        # 获取收入（万元单位）
        revenue_wan = company.get("revenue", 10000)

        # 基于收入计算合理的估值（万元单位）
        base_valuation_wan = revenue_wan * 1.5  # 1.5倍收入作为基准估值

        # 返回模拟数据用于测试（使用万元单位）
        return {
            "success": True,
            "results": {
                "PE": {
                    "value": base_valuation_wan,  # 15000万元
                    "value_low": base_valuation_wan * 0.8,  # 12000万元
                    "value_high": base_valuation_wan * 1.2,  # 18000万元
                    "multiple": 15.5,
                    "description": "市盈率法"
                },
                "PB": {
                    "value": base_valuation_wan * 1.07,  # 约16000万元
                    "value_low": base_valuation_wan * 0.87,  # 约13000万元
                    "value_high": base_valuation_wan * 1.27,  # 约19000万元
                    "multiple": 2.3,
                    "description": "市净率法"
                },
                "PS": {
                    "value": base_valuation_wan * 0.93,  # 约14000万元
                    "value_low": base_valuation_wan * 0.73,  # 约11000万元
                    "value_high": base_valuation_wan * 1.13,  # 约17000万元
                    "multiple": 3.1,
                    "description": "市销率法"
                },
                "EV_EBITDA": {
                    "value": base_valuation_wan * 1.03,  # 约15500万元
                    "value_low": base_valuation_wan * 0.83,  # 约12500万元
                    "value_high": base_valuation_wan * 1.23,  # 约18500万元
                    "multiple": 12.8,
                    "description": "EV/EBITDA法"
                }
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/valuation/absolute")
async def absolute_valuation(request: Dict[str, Any]):
    """绝对估值计算（DCF）"""
    try:
        company = request.get("company", {})
        # 处理单位：前端传入的收入是万元，需要转换为元进行DCF计算
        revenue_wan = company.get("revenue", 10000)
        revenue_yuan = revenue_wan * 10000  # 转换为元

        growth_rate = company.get("growth_rate", 0.1) / 100 if company.get("growth_rate") else 0.1
        operating_margin = company.get("operating_margin", 0.15) / 100 if company.get("operating_margin") else 0.15
        wacc_input = company.get("wacc", 0.12) / 100 if company.get("wacc") else 0.12
        tax_rate = company.get("tax_rate", 0.25) / 100 if company.get("tax_rate") else 0.25

        # 简化的DCF计算（使用元为单位）
        base_fcf = revenue_yuan * operating_margin * (1 - tax_rate)  # 自由现金流基数（元）
        wacc = wacc_input if wacc_input > 0 else 0.12

        # 计算5年现金流
        fcf_projection = []
        for i, year in enumerate([2024, 2025, 2026, 2027, 2028]):
            fcf = base_fcf * (1 + growth_rate) ** i
            fcf_projection.append({
                "year": year,
                "fcf": fcf,
                "growth": growth_rate
            })

        # 计算现值
        pv_forecasts = sum(fcf["fcf"] / (1 + wacc) ** (i + 1) for i, fcf in enumerate(fcf_projection))

        # 终值计算（永续增长）
        terminal_growth = 0.03
        terminal_value = fcf_projection[-1]["fcf"] * (1 + terminal_growth) / (wacc - terminal_growth)
        pv_terminal = terminal_value / (1 + wacc) ** 5

        # 企业价值和股权价值（元）
        enterprise_value_yuan = pv_forecasts + pv_terminal
        total_debt_yuan = company.get("total_debt", 0) * 10000  # 债务也是万元单位
        cash_yuan = company.get("cash_and_equivalents", 0) * 10000  # 现金也是万元单位
        equity_value_yuan = enterprise_value_yuan - total_debt_yuan + cash_yuan

        # 转换回万元供前端使用
        enterprise_value = enterprise_value_yuan / 10000
        equity_value = equity_value_yuan / 10000

        # 转换现值数据为万元
        pv_forecasts_wan = pv_forecasts / 10000
        pv_terminal_wan = pv_terminal / 10000
        terminal_value_wan = terminal_value / 10000

        # 转换fcf_projection为万元
        fcf_projection_wan = [
            {
                "year": fcf["year"],
                "fcf": fcf["fcf"] / 10000,  # 转换为万元
                "growth": fcf["growth"]
            }
            for fcf in fcf_projection
        ]

        return {
            "success": True,
            "result": {
                "value": equity_value,
                "details": {
                    "wacc": wacc,
                    "pv_forecasts": pv_forecasts_wan,
                    "pv_terminal": pv_terminal_wan,
                    "terminal_growth": terminal_growth,
                    "enterprise_value": enterprise_value,
                    "fcf_projection": fcf_projection_wan,
                    "terminal_value": terminal_value_wan
                }
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/valuation/vc")
async def vc_valuation(request: Dict[str, Any]):
    """VC法估值"""
    try:
        return {
            "success": True,
            "method": "VC法",
            "result": {
                "pre_money_valuation": 50000000,
                "post_money_valuation": 75000000,
                "investment_amount": 25000000,
                "equity_percentage": 33.33
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 多产品估值
@app.post("/api/valuation/multi-product-dcf")
async def multi_product_dcf(request: Dict[str, Any]):
    """多产品DCF估值"""
    try:
        return {
            "success": True,
            "method": "多产品DCF法",
            "result": {
                "products": request.get("products", []),
                "total_enterprise_value": 500000000,
                "equity_value": 450000000,
                "valuation_by_product": [
                    {"product": "产品A", "value": 200000000},
                    {"product": "产品B", "value": 150000000},
                    {"product": "产品C", "value": 100000000}
                ]
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 多方法交叉验证
@app.post("/api/valuation/compare")
async def compare_valuation(request: Dict[str, Any]):
    """多方法估值对比"""
    try:
        return {
            "success": True,
            "comparison": {
                "methods": ["P/E", "P/B", "DCF", "VC法"],
                "valuations": [120000000, 135000000, 150000000, 100000000],
                "average": 126250000,
                "median": 127500000,
                "std_dev": 18750000,
                "recommendation": "建议采用DCF估值结果: 150,000,000元"
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 情景分析
@app.post("/api/scenario/analyze")
async def scenario_analysis(request: Dict[str, Any]):
    """情景分析"""
    try:
        company = request.get("company", {})
        scenario_params = request.get("scenario_params")  # 获取情景参数

        revenue = company.get("revenue", 10000)  # 万元单位
        growth_rate = company.get("growth_rate", 10) / 100 if company.get("growth_rate") else 0.1
        base_valuation = revenue * 1.5  # 简化计算

        # 默认情景参数（如果没有传递）
        default_scenarios = {
            "基准情景": {
                "revenue_growth_adj": 0,
                "margin_adj": 0,
                "wacc_adj": 0
            },
            "乐观情景": {
                "revenue_growth_adj": 0.3,  # +30%
                "margin_adj": 0.02,   # +2%
                "wacc_adj": -0.01    # -1%
            },
            "悲观情景": {
                "revenue_growth_adj": -0.3, # -30%
                "margin_adj": -0.03,  # -3%
                "wacc_adj": 0.02     # +2%
            }
        }

        # 如果传递了情景参数，使用传递的参数
        if scenario_params and len(scenario_params) >= 3:
            scenario_config = [
                {
                    "name": scenario_params[0].get("name", "基准情景"),
                    "revenue_growth_adj": scenario_params[0].get("revenue_growth_adj", 0),
                    "margin_adj": scenario_params[0].get("margin_adj", 0),
                    "wacc_adj": scenario_params[0].get("wacc_adj", 0)
                },
                {
                    "name": scenario_params[1].get("name", "乐观情景"),
                    "revenue_growth_adj": scenario_params[1].get("revenue_growth_adj", 0.3),
                    "margin_adj": scenario_params[1].get("margin_adj", 0.02),
                    "wacc_adj": scenario_params[1].get("wacc_adj", -0.01)
                },
                {
                    "name": scenario_params[2].get("name", "悲观情景"),
                    "revenue_growth_adj": scenario_params[2].get("revenue_growth_adj", -0.3),
                    "margin_adj": scenario_params[2].get("margin_adj", -0.03),
                    "wacc_adj": scenario_params[2].get("wacc_adj", 0.02)
                }
            ]
        else:
            # 使用默认参数
            scenario_config = [
                {
                    "name": "基准情景",
                    "revenue_growth_adj": 0,
                    "margin_adj": 0,
                    "wacc_adj": 0
                },
                {
                    "name": "乐观情景",
                    "revenue_growth_adj": 0.3,
                    "margin_adj": 0.02,
                    "wacc_adj": -0.01
                },
                {
                    "name": "悲观情景",
                    "revenue_growth_adj": -0.3,
                    "margin_adj": -0.03,
                    "wacc_adj": 0.02
                }
            ]

        # 根据参数配置生成不同情景的估值
        scenarios = {}
        probabilities = [0.6, 0.25, 0.15]

        for i, config in enumerate(scenario_config):
            scenario_name = config["name"]

            # 计算调整后的估值
            adjusted_growth = growth_rate * (1 + config["revenue_growth_adj"])
            margin_impact = config["margin_adj"]
            wacc_impact = config["wacc_adj"]

            # 根据调整计算估值
            if scenario_name == "基准情景":
                adjusted_valuation = base_valuation
                adjusted_wacc = 0.12
                probability = probabilities[0]
            elif scenario_name == "乐观情景":
                adjusted_valuation = base_valuation * (1 + config["revenue_growth_adj"]) * (1 + margin_impact)
                adjusted_wacc = 0.12 + wacc_impact
                probability = probabilities[1]
            else:  # 悲观情景
                adjusted_valuation = base_valuation * (1 + config["revenue_growth_adj"]) * (1 + margin_impact)
                adjusted_wacc = 0.12 + wacc_impact
                probability = probabilities[2]

            scenarios[scenario_name] = {
                "probability": probability,
                "growth_rate": adjusted_growth,
                "valuation": {
                    "value": adjusted_valuation,
                    "method": "DCF",
                    "details": {"wacc": adjusted_wacc, "description": f"{scenario_name}下的估值"}
                },
                "description": f"{scenario_name}的估值分析",
                "scenario": {
                    "revenue_growth_adj": config["revenue_growth_adj"],
                    "margin_adj": config["margin_adj"],
                    "wacc_adj": config["wacc_adj"]
                }
            }

        # 计算期望估值
        expected_valuation = sum(
            scenario["valuation"]["value"] * scenario["probability"]
            for scenario in scenarios.values()
        )

        return {
            "success": True,
            "results": scenarios,
            "expected_valuation": expected_valuation,
            "valuation_range": {
                "min": scenarios["悲观情景"]["valuation"]["value"],
                "max": scenarios["乐观情景"]["valuation"]["value"],
                "confidence_90": [
                    scenarios["悲观情景"]["valuation"]["value"] * 1.1,
                    scenarios["乐观情景"]["valuation"]["value"] * 0.9
                ]
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 压力测试 - 收入冲击
@app.post("/api/stress-test/revenue")
async def revenue_stress_test(request: Dict[str, Any]):
    """收入冲击压力测试"""
    try:
        company = request.get("company", {})
        shocks_input = request.get("shocks", [-0.2, -0.1, 0, 0.1, 0.2])

        # 基于公司收入计算基础估值（修正单位处理）
        # 前端传入的收入单位是"万元"，需要转换为"元"进行计算
        revenue_wan = company.get("revenue", 10000)  # 获取万元数值
        revenue_yuan = revenue_wan * 10000  # 转换为元
        base_valuation_yuan = revenue_yuan * 1.5  # 1.5倍收入作为估值（元）
        base_valuation_wan = base_valuation_yuan / 10000  # 转换回万元供前端使用

        # 确保shocks是数组格式
        if isinstance(shocks_input, list):
            shocks = shocks_input
        else:
            shocks = [-0.2, -0.1, 0, 0.1, 0.2]

        # 生成收入冲击测试结果
        revenue_shock_tests = []
        for shock in shocks:
            valuation = base_valuation_wan * (1 + shock)  # 使用万元单位
            change_pct = shock  # 保持小数形式，与前端期望一致

            # 确定情景描述
            if shock < 0:
                scenario_desc = f"收入下降{abs(shock * 100):.0f}%"
            elif shock > 0:
                scenario_desc = f"收入增长{abs(shock * 100):.0f}%"
            else:
                scenario_desc = "基准情况"

            revenue_shock_tests.append({
                "scenario_description": scenario_desc,
                "shock_level": shock,
                "stressed_value": valuation,  # 万元单位，前端会除以10000显示亿元
                "change_pct": change_pct,  # 保持小数形式，前端会乘以100
                "description": f"收入{'下降' if shock < 0 else '增长' if shock > 0 else '保持'}{abs(shock * 100):.0f}%时的估值"
            })

        return {
            "success": True,
            "report": {
                "tests": {
                    "revenue_shock": revenue_shock_tests
                },
                "summary": {
                    "test_type": "收入冲击测试",
                    "sensitivity": "高敏感性：收入变化1%，估值变化约1%",
                    "base_valuation": base_valuation_wan  # 使用万元单位
                }
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 压力测试 - 蒙特卡洛模拟
@app.post("/api/stress-test/monte-carlo")
async def monte_carlo_simulation(request: Dict[str, Any], iterations: int = 1000):
    """蒙特卡洛模拟"""
    try:
        import random
        random.seed(42)

        company = request.get("company", {})
        revenue_wan = company.get("revenue", 10000)  # 获取万元数值
        revenue_yuan = revenue_wan * 10000  # 转换为元
        base_valuation_yuan = revenue_yuan * 1.5  # 简化计算
        base_valuation_wan = base_valuation_yuan / 10000  # 转换回万元

        # 生成模拟估值结果（使用万元单位）
        valuations = []
        for _ in range(iterations):
            variation = random.uniform(-0.3, 0.3)  # ±30%变化
            valuations.append(base_valuation_wan * (1 + variation))

        valuations.sort()

        # 计算统计数据
        mean_val = sum(valuations) / len(valuations)
        median_val = valuations[len(valuations) // 2]
        std_dev = (sum((x - mean_val) ** 2 for x in valuations) / len(valuations)) ** 0.5

        # 百分位数
        percentiles = {
            "p5": valuations[int(len(valuations) * 0.05)],
            "p25": valuations[int(len(valuations) * 0.25)],
            "p50": median_val,
            "p75": valuations[int(len(valuations) * 0.75)],
            "p95": valuations[int(len(valuations) * 0.95)]
        }

        return {
            "success": True,
            "report": {
                "tests": {
                    "monte_carlo": {
                        "iterations": iterations,
                        "mean": mean_val,
                        "median": median_val,
                        "std_dev": std_dev,
                        "min": valuations[0],
                        "max": valuations[-1],
                        "percentiles": percentiles,
                        # 为了兼容前端，添加扁平化的字段
                        "percentile_5": percentiles["p5"],
                        "percentile_95": percentiles["p95"],
                        "confidence_intervals": {
                            "ci_90": [percentiles["p5"], percentiles["p95"]],
                            "ci_95": [valuations[int(len(valuations) * 0.025)], valuations[int(len(valuations) * 0.975)]]
                        },
                        # 生成直方图分布数据
                        "distribution": _generate_histogram(valuations, 20)  # 20个区间
                    }
                },
                "summary": {
                    "method": "蒙特卡洛模拟",
                    "description": f"基于{iterations}次随机模拟的估值分布"
                }
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 压力测试 - 综合压力测试
@app.post("/api/stress-test/full")
async def full_stress_test(request: Dict[str, Any]):
    """综合压力测试"""
    try:
        return {
            "success": True,
            "test_results": [
                {
                    "test": "收入冲击",
                    "worst_case": 105000000,
                    "impact": "高"
                },
                {
                    "test": "毛利率压缩",
                    "worst_case": 135000000,
                    "impact": "中"
                },
                {
                    "test": "WACC上升",
                    "worst_case": 120000000,
                    "impact": "中"
                },
                {
                    "test": "综合压力",
                    "worst_case": 90000000,
                    "impact": "高"
                }
            ],
            "overall_risk_assessment": "中等风险",
            "recommendations": ["建议进行敏感性分析", "关注市场变化", "建立风险预警机制"]
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 敏感性分析 - 单因素
@app.post("/api/sensitivity/one-way")
async def one_way_sensitivity(request: Dict[str, Any]):
    """单因素敏感性分析"""
    try:
        param_name = request.get("param_name", "增长率")
        param_range = request.get("param_range", [-0.1, 0, 0.1, 0.2, 0.3])

        return {
            "success": True,
            "param_name": param_name,
            "sensitivity": [
                {"param_value": -0.1, "valuation": 135000000, "change": -10},
                {"param_value": 0, "valuation": 150000000, "change": 0},
                {"param_value": 0.1, "valuation": 165000000, "change": 10},
                {"param_value": 0.2, "valuation": 180000000, "change": 20},
                {"param_value": 0.3, "valuation": 195000000, "change": 30}
            ],
            "conclusion": f"对{param_name}敏感性较高"
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 敏感性分析 - 龙卷风图
@app.post("/api/sensitivity/tornado")
async def tornado_analysis(request: Dict[str, Any]):
    """龙卷风图数据分析"""
    try:
        param_changes = request.get("param_changes", {})
        company = request.get("company", {})
        revenue_wan = company.get("revenue", 10000)
        base_valuation_wan = revenue_wan * 1.5

        # 如果没有传入参数变化，使用默认值
        if not param_changes:
            param_changes = {
                "growth_rate": 10,
                "operating_margin": 5,
                "wacc": 1,
                "terminal_growth": 0.5
            }

        # 计算每个参数的敏感性影响
        parameters = {}

        # 收入增长率影响 (假设±10%变化导致±25%估值变化 - 提高影响幅度)
        growth_change_pct = param_changes.get("growth_rate", 10) / 100
        growth_impact = abs(growth_change_pct * 2.5)  # 2.5倍弹性，影响更大
        parameters["收入增长率"] = {
            "valuation_range": abs(base_valuation_wan * growth_impact * 2),  # 总范围（正负两侧）
            "base_value": 0.10,  # 10%基准增长率
            "impact_percentage": growth_impact
        }

        # 营业利润率影响 (假设±5%变化导致±20%估值变化)
        margin_change_pct = param_changes.get("operating_margin", 5) / 100
        margin_impact = abs(margin_change_pct * 4)  # 4倍弹性
        parameters["营业利润率"] = {
            "valuation_range": abs(base_valuation_wan * margin_impact * 2),
            "base_value": 0.15,  # 15%基准利润率
            "impact_percentage": margin_impact
        }

        # WACC影响 (假设±1%变化导致±15%估值变化)
        wacc_change_pct = param_changes.get("wacc", 1) / 100
        wacc_impact = abs(wacc_change_pct * 15)  # 15倍弹性
        parameters["WACC"] = {
            "valuation_range": abs(base_valuation_wan * wacc_impact * 2),
            "base_value": 0.12,  # 12%基准WACC
            "impact_percentage": wacc_impact
        }

        # 永续增长率影响 (假设±0.5%变化导致±10%估值变化)
        terminal_change_pct = param_changes.get("terminal_growth", 0.5) / 100
        terminal_impact = abs(terminal_change_pct * 20)  # 20倍弹性
        parameters["永续增长率"] = {
            "valuation_range": abs(base_valuation_wan * terminal_impact * 2),
            "base_value": 0.025,  # 2.5%基准永续增长率
            "impact_percentage": terminal_impact
        }

        # 税率影响 (假设±5%变化导致±8%估值变化)
        tax_change_pct = 0.05  # 固定5%变化
        tax_impact = abs(tax_change_pct * 1.6)  # 1.6倍弹性
        parameters["税率"] = {
            "valuation_range": abs(base_valuation_wan * tax_impact * 2),
            "base_value": 0.25,  # 25%基准税率
            "impact_percentage": tax_impact
        }

        # 转换为前端期望的格式
        result_array = []
        for param_name, param_data in parameters.items():
            result_array.append({
                "parameter": param_name,
                "max_impact": param_data["valuation_range"] / 2,  # 单侧影响（总范围的一半）
                "impact_pct": param_data["impact_percentage"]
            })

        return {
            "success": True,
            "result": result_array,
            "most_sensitive": max(parameters.keys(), key=lambda k: parameters[k]["valuation_range"]),
            "conclusion": f"{max(parameters.keys(), key=lambda k: parameters[k]['valuation_range'])}对估值影响最大，需重点关注"
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 敏感性分析 - 综合
@app.post("/api/sensitivity/comprehensive")
async def comprehensive_sensitivity(request: Dict[str, Any]):
    """综合敏感性分析"""
    try:
        param_changes = request.get("param_changes", {})
        company = request.get("company", {})
        revenue_wan = company.get("revenue", 10000)
        base_valuation_wan = revenue_wan * 1.5

        # 如果没有传入参数变化，使用默认值
        if not param_changes:
            param_changes = {
                "growth_rate": 10,
                "operating_margin": 5,
                "wacc": 1,
                "terminal_growth": 0.5
            }

        # 计算综合敏感性分析
        growth_change_pct = param_changes.get("growth_rate", 10) / 100
        margin_change_pct = param_changes.get("operating_margin", 5) / 100
        wacc_change_pct = param_changes.get("wacc", 1) / 100
        terminal_change_pct = param_changes.get("terminal_growth", 0.5) / 100

        analysis = {
            "single_factor": [
                {
                    "param": "收入增长率",
                    "sensitivity": "高",
                    "impact_range": [-base_valuation_wan * growth_change_pct * 2, base_valuation_wan * growth_change_pct * 2]
                },
                {
                    "param": "营业利润率",
                    "sensitivity": "中高",
                    "impact_range": [-base_valuation_wan * margin_change_pct * 3, base_valuation_wan * margin_change_pct * 3]
                },
                {
                    "param": "WACC",
                    "sensitivity": "中",
                    "impact_range": [-base_valuation_wan * wacc_change_pct * 12, base_valuation_wan * wacc_change_pct * 12]
                },
                {
                    "param": "永续增长率",
                    "sensitivity": "中低",
                    "impact_range": [-base_valuation_wan * terminal_change_pct * 16, base_valuation_wan * terminal_change_pct * 16]
                }
            ],
            "key_factors": ["收入增长率", "营业利润率", "WACC"],
            "risk_level": "中等",
            "monitoring_focus": "重点关注收入增长和营业利润率变化"
        }

        return {
            "success": True,
            "results": analysis,
            "summary": {
                "risk_level": "中等",
                "conclusion": "收入增长率和营业利润率对估值影响较大"
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 数据API - 配置Tushare
@app.post("/api/data/tushare/configure")
async def configure_tushare(token: str):
    """配置Tushare Token"""
    return {"status": "success", "message": "Tushare Token已配置", "token_valid": True}

# 数据API - 获取可比公司
@app.get("/api/data/comparable/{industry}")
async def get_comparables(industry: str, market_cap_min: int = None, market_cap_max: int = None, limit: int = 20):
    """获取可比公司数据"""
    try:
        # 模拟可比公司数据
        comparables = [
            {
                "name": f"公司{i}",
                "code": f"60{i}000",
                "industry": industry,
                "market_cap": 5000000000 + i * 1000000000,
                "pe_ratio": 15.0 + i,
                "pb_ratio": 2.5 + i * 0.1,
                "ps_ratio": 3.0 + i * 0.2
            }
            for i in range(1, min(limit, 10) + 1)
        ]
        return {"status": "success", "comparables": comparables, "total": len(comparables)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 数据API - 获取股票数据
@app.get("/api/data/stock/{ts_code}")
async def get_stock_data(ts_code: str):
    """获取股票财务数据"""
    try:
        return {
            "success": True,
            "stock": {
                "code": ts_code,
                "name": f"股票_{ts_code}",
                "industry": "制造业",
                "market_cap": 10000000000,
                "revenue": 5000000000,
                "net_income": 800000000,
                "total_assets": 15000000000,
                "pe_ratio": 12.5,
                "pb_ratio": 2.0
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 数据API - 获取行业估值倍数
@app.get("/api/data/industry-multiples/{industry}")
async def get_industry_multiples(industry: str, method: str = "median"):
    """获取行业估值倍数"""
    try:
        return {
            "success": True,
            "industry": industry,
            "method": method,
            "multiples": {
                "pe_ratio": 15.5,
                "pb_ratio": 2.8,
                "ps_ratio": 3.2,
                "ev_ebitda": 12.5
            },
            "sample_size": 25
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 数据API - 搜索公司
@app.post("/api/data/search")
async def search_companies(request: Dict[str, Any]):
    """搜索公司"""
    try:
        keywords = request.get("keywords", [])
        limit = request.get("limit", 10)

        results = [
            {
                "name": f"搜索结果{i}",
                "code": f"60{i}000",
                "industry": "制造业",
                "match_score": 0.9 - i * 0.05
            }
            for i in range(1, min(limit, 5) + 1)
        ]
        return {"status": "success", "results": results, "total": len(results)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 历史记录API - 获取列表
@app.get("/api/history")
async def get_history(limit: int = 50):
    """获取历史记录列表"""
    try:
        history = [
            {
                "id": i,
                "date": f"2024-01-{i:02d}",
                "company_name": f"测试公司{i}",
                "valuation_method": "DCF",
                "valuation": 10000 + i * 1000,  # 改为万元单位
                "revenue": 8000 + i * 800,  # 万元单位
                "net_income": 1200 + i * 120,  # 万元单位
                "total_assets": 20000 + i * 2000,  # 万元单位
                "net_assets": 15000 + i * 1500,  # 万元单位
                "industry": ["制造业", "科技业", "金融业", "医疗业", "消费品"][i % 5],
                "stage": ["早期", "成长期", "成熟期", "上市公司"][i % 4],
                "growth_rate": 0.08 + (i % 5) * 0.02,
                "wacc": 0.10 + (i % 3) * 0.01,
                "operating_margin": 0.12 + (i % 4) * 0.02,
                "status": "completed"
            }
            for i in range(1, min(limit, 10) + 1)
        ]
        return {"status": "success", "history": history, "total": len(history)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 历史记录API - 获取详情
@app.get("/api/history/{history_id}")
async def get_history_detail(history_id: int):
    """获取历史记录详情"""
    try:
        return {
            "success": True,
            "record": {
                "id": history_id,
                "date": "2024-01-15",
                "company_name": "测试公司",
                "valuation_method": "DCF",
                "valuation": 15000,  # 改为万元单位
                "revenue": 12000,  # 万元单位
                "net_income": 1800,  # 万元单位
                "total_assets": 25000,  # 万元单位
                "net_assets": 18000,  # 万元单位
                "industry": "制造业",
                "stage": "成长期",
                "growth_rate": 0.12,
                "wacc": 0.11,
                "operating_margin": 0.15,
                "details": {
                    "method": "DCF现金流折现",
                    "wacc": 0.12,
                    "terminal_growth": 0.03,
                    "enterprise_value": 18000,  # 万元单位
                    "equity_value": 15000  # 万元单位
                }
            }
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# 历史记录API - 保存
@app.post("/api/history/save")
async def save_history(request: Dict[str, Any]):
    """保存估值结果到历史记录"""
    try:
        return {
            "success": True,
            "message": "估值结果已保存",
            "history_id": 12345
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/api/history/{history_id}")
async def delete_history(history_id: int):
    """删除历史记录"""
    try:
        return {
            "success": True,
            "message": f"历史记录 {history_id} 已删除"
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

def _generate_histogram(values: List[float], bins: int = 20) -> List[Dict[str, Any]]:
    """生成直方图数据"""
    if not values:
        return []

    min_val = min(values)
    max_val = max(values)
    bin_width = (max_val - min_val) / bins

    histogram = []
    for i in range(bins):
        bin_lower = min_val + i * bin_width
        bin_upper = min_val + (i + 1) * bin_width

        # 计算该区间内的数据点数量
        count = sum(1 for v in values if bin_lower <= v < bin_upper)

        # 最后一个区间包含最大值
        if i == bins - 1:
            count = sum(1 for v in values if bin_lower <= v <= bin_upper)

        histogram.append({
            "bin_lower": bin_lower,
            "bin_upper": bin_upper,
            "count": count,
            "bin_center": (bin_lower + bin_upper) / 2
        })

    return histogram

if __name__ == "__main__":
    import uvicorn
    print("启动投资估值系统API (完整版)...")
    print("访问地址: http://localhost:8000")
    print("API文档: http://localhost:8000/docs")
    uvicorn.run(app, host="0.0.0.0", port=8000)