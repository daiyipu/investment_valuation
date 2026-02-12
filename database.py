"""
数据库模型 - 估值历史记录
"""
from datetime import datetime
from sqlalchemy import Column, Integer, String, DateTime, Float, Text
from sqlalchemy.ext.declarative import declarative_base
from pydantic import BaseModel


# 估值历史记录表
class ValuationHistory(BaseModel):
    """估值历史记录"""
    __tablename__ = 'valuation_history'

    id = Column(Integer, primary_key=True, index=True, comment='记录ID')

    # 公司基本信息
    company_name = Column(String(100), comment='公司名称', index=True)
    industry = Column(String(50), comment='所属行业', index=True)
    stage = Column(String(20), comment='发展阶段', index=True)

    # 财务数据
    revenue = Column(Float, comment='营业收入（万元）', index=True)
    net_income = Column(Float, comment='净利润（万元）', index=True)
    net_assets = Column(Float, comment='净资产（万元）', index=True)
    ebitda = Column(Float, comment='EBITDA（万元）', index=True)

    # 估值参数
    growth_rate = Column(Float, comment='预期增长率（%）', index=True)
    operating_margin = Column(Float, comment='营业利润率（%）', index=True)
    beta = Column(Float, comment='贝塔系数', index=True)
    risk_free_rate = Column(Float, comment='无风险利率', index=True)
    market_risk_premium = Column(Float, comment='市场风险溢价', index=True)
    terminal_growth_rate = Column(Float, comment='永续增长率（%）', index=True)

    # 估值结果
    dcf_value = Column(Float, comment='DCF估值（万元）', index=True)
    dcf_wacc = Column(Float, comment='DCF-WACC（%）', index=True)

    # 相对估值结果
    pe_value = Column(Float, comment='P/E估值（万元）', index=True)
    pe_ratio = Column(Float, comment='P/E倍数', index=True)
    ps_value = Column(Float, comment='P/S估值（万元）', index=True)
    ps_ratio = Column(Float, comment='P/S倍数', index=True)
    pb_value = Column(Float, comment='P/B估值（万元）', index=True)
    ev_value = Column(Float, comment='EV/EBITDA估值（万元）', index=True)
    ev_ebitda_ratio = Column(Float, comment='EV/EBITDA倍数', index=True)

    # 可比公司数量
    comparables_count = Column(Integer, comment='可比公司数量', index=True)

    # 创建时间
    created_at = Column(DateTime, default=datetime.utcnow, comment='创建时间', index=True)

    # 备注
    notes = Column(Text, comment='备注', index=True)


# 数据库操作基类
class DatabaseManager:
    """数据库管理器"""

    def __init__(self, database_url: str = 'sqlite:///investment_valuation.db'):
        self.database_url = database_url
        self.engine = None

    def get_engine(self):
        """获取数据库引擎"""
        if self.engine is None:
            from sqlalchemy import create_engine
            self.engine = create_engine(
                self.database_url,
                connect_args={'check_same_thread': False},
                echo=False
            )
        return self.engine

    def create_tables(self):
        """创建所有表"""
        from sqlalchemy import create_all_tables
        Base.metadata.create_all(self.engine)

    def init_history(self, company_data: dict, results: dict) -> int:
        """初始化估值历史记录"""
        session = Session(self.engine)

        # 创建历史记录
        history = ValuationHistory(
            company_name=company_data.get('name', ''),
            industry=company_data.get('industry', ''),
            stage=company_data.get('stage', ''),
            revenue=float(company_data.get('revenue', 0)) / 10000,  # 转换为万元
            net_income=float(company_data.get('net_income', 0)) / 10000,
            net_assets=float(company_data.get('net_assets', 0)) / 10000,
            ebitda=float(company_data.get('ebitda', 0)) / 10000,
            growth_rate=float(company_data.get('growth_rate', 0)),
            operating_margin=float(company_data.get('operating_margin', 0)),
            beta=float(company_data.get('beta', 0)),
            risk_free_rate=float(company_data.get('risk_free_rate', 0)),
            market_risk_premium=float(company_data.get('market_risk_premium', 0)),
            terminal_growth_rate=float(company_data.get('terminal_growth_rate', 0)),
            # DCF结果
            dcf_value=float(results.get('dcf', {}).get('value', 0)) / 10000,
            dcf_wacc=float(results.get('dcf', {}).get('details', {}).get('wacc', 0)),
            # 相对估值结果
            pe_value=float(results.get('relative', {}).get('result', {}).get('pe_valuation', 0)) / 10000 if results.get('relative', {}).get('result', {}).get('pe_ratio', 0) else None,
            ps_value=float(results.get('relative', {}).get('result', {}).get('ps_valuation', 0)) / 10000 if results.get('relative', {}).get('result', {}).get('ps_ratio', 0) else None,
            pb_value=float(results.get('relative', {}).get('result', {}).get('pb_valuation', 0)) / 10000 if results.get('relative', {}).get('result', {}).get('pb_ratio', 0) else None,
            ev_value=float(results.get('relative', {}).get('result', {}).get('ev_valuation', 0)) / 10000 if results.get('relative', {}).get('result', {}).get('ev_ebitda_ratio', 0) else None,
            ev_ebitda_ratio=results.get('relative', {}).get('ev_ebitda', 0) if results.get('relative', {}).get('ev_ebitda_ratio', 0) else None,
            # 可比公司数量
            comparables_count=len(results.get('relative', {}).get('comparables', [])),
        )

        session.commit()
        session.flush()
        session.refresh(history)

        return history.id
