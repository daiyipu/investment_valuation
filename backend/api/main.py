"""
投资估值系统 - FastAPI后端服务
提供RESTful API供前端调用
"""
from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Optional, Dict, Any
import os
import sys

# 添加项目根目录到 Python 路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from core.models import Company, Comparable, CompanyStage, ValuationResult, ScenarioConfig, ProductSegment
from core.database import DatabaseManager
from services.relative_valuation import RelativeValuation
from services.absolute_valuation import AbsoluteValuation
from utils.other_methods import OtherValuationMethods
from services.scenario_analysis import ScenarioAnalyzer
from services.stress_test import StressTester
from services.sensitivity_analysis import SensitivityAnalyzer
from services.multi_product_valuation import MultiProductValuation, validate_products
from api.schemas import (
    CompanyInput, ComparableInput, ScenarioInput,
    RelativeValuationRequest, ProductSegmentInput,
    MultiProductValuationRequest, ScenarioAnalysisRequest
)

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


# ===== 导入所有路由 =====
from api.routes.valuation import router as valuation_router
from api.routes.scenario import router as scenario_router
from api.routes.stress_test import router as stress_test_router
from api.routes.sensitivity import router as sensitivity_router
from api.routes.data import router as data_router
from api.routes.history import router as history_router

# 注册路由
app.include_router(valuation_router, tags=["估值"])
app.include_router(scenario_router, prefix="/scenario", tags=["情景分析"])
app.include_router(stress_test_router, prefix="/stress-test", tags=["压力测试"])
app.include_router(sensitivity_router, prefix="/sensitivity", tags=["敏感性分析"])
app.include_router(data_router, prefix="/data", tags=["数据获取"])
app.include_router(history_router, prefix="/history", tags=["历史记录"])


@app.get("/")
async def root():
    """根路径"""
    return {
        "message": "投资估值系统API",
        "version": "1.0.0",
        "docs": "/docs"
    }


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
