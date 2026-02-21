"""
数据库模型 - 估值历史记录
"""
from datetime import datetime
from sqlalchemy import Column, Integer, String, DateTime, Float, Text, create_engine
from sqlalchemy.orm import declarative_base, Session
from pydantic import BaseModel
from typing import Optional, List


# SQLAlchemy 声明式基类
Base = declarative_base()


# 估值历史记录表
class ValuationHistory(Base):
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

    def to_dict(self):
        """转换为字典"""
        return {
            'id': self.id,
            'company_name': self.company_name,
            'industry': self.industry,
            'stage': self.stage,
            'revenue': self.revenue,
            'net_income': self.net_income,
            'net_assets': self.net_assets,
            'ebitda': self.ebitda,
            'growth_rate': self.growth_rate,
            'operating_margin': self.operating_margin,
            'beta': self.beta,
            'risk_free_rate': self.risk_free_rate,
            'market_risk_premium': self.market_risk_premium,
            'terminal_growth_rate': self.terminal_growth_rate,
            'dcf_value': self.dcf_value,
            'dcf_wacc': self.dcf_wacc,
            'pe_value': self.pe_value,
            'pe_ratio': self.pe_ratio,
            'ps_value': self.ps_value,
            'ps_ratio': self.ps_ratio,
            'pb_value': self.pb_value,
            'ev_value': self.ev_value,
            'ev_ebitda_ratio': self.ev_ebitda_ratio,
            'comparables_count': self.comparables_count,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'notes': self.notes,
        }


# Pydantic 模型用于API响应
class ValuationHistoryResponse(BaseModel):
    """估值历史记录响应模型"""
    id: int
    company_name: str
    industry: Optional[str] = None
    stage: Optional[str] = None
    revenue: Optional[float] = None
    dcf_value: Optional[float] = None
    dcf_wacc: Optional[float] = None
    pe_value: Optional[float] = None
    ps_value: Optional[float] = None
    pb_value: Optional[float] = None
    ev_value: Optional[float] = None
    comparables_count: Optional[int] = None
    created_at: Optional[str] = None


# 数据库操作类
class DatabaseManager:
    """数据库管理器"""

    def __init__(self, database_url: str = 'sqlite:///investment_valuation.db'):
        self.database_url = database_url
        self.engine = None

    def get_engine(self):
        """获取数据库引擎"""
        if self.engine is None:
            self.engine = create_engine(
                self.database_url,
                connect_args={'check_same_thread': False},
                echo=False
            )
        return self.engine

    def create_tables(self):
        """创建所有表"""
        engine = self.get_engine()
        Base.metadata.create_all(engine)

    def init_history(self, company_data: dict, results: dict) -> int:
        """初始化估值历史记录"""
        session = Session(self.get_engine())

        try:
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
                dcf_value=float(results.get('dcf', {}).get('result', {}).get('value', 0)) / 10000,
                dcf_wacc=float(results.get('dcf', {}).get('result', {}).get('details', {}).get('wacc', 0)),
                # 相对估值结果
                pe_value=self._get_relative_value(results, 'pe') / 10000 if self._has_relative_method(results, 'pe') else None,
                ps_value=self._get_relative_value(results, 'ps') / 10000 if self._has_relative_method(results, 'ps') else None,
                pb_value=self._get_relative_value(results, 'pb') / 10000 if self._has_relative_method(results, 'pb') else None,
                ev_value=self._get_relative_value(results, 'ev') / 10000 if self._has_relative_method(results, 'ev') else None,
                ev_ebitda_ratio=results.get('relative', {}).get('result', {}).get('ev_ebitda', 0) if self._has_relative_method(results, 'ev') else None,
                # 可比公司数量
                comparables_count=len(results.get('relative', {}).get('comparables', [])),
            )

            session.add(history)
            session.commit()
            session.flush()
            session.refresh(history)

            return history.id
        finally:
            session.close()

    def _get_relative_value(self, results: dict, method: str) -> float:
        """获取相对估值方法的结果值"""
        rel_result = results.get('relative', {}).get('result', {})
        if method == 'pe':
            return rel_result.get('pe_valuation', 0)
        elif method == 'ps':
            return rel_result.get('ps_valuation', 0)
        elif method == 'pb':
            return rel_result.get('pb_valuation', 0)
        elif method == 'ev':
            return rel_result.get('ev_valuation', 0)
        return 0

    def _has_relative_method(self, results: dict, method: str) -> bool:
        """检查是否有某种相对估值方法的结果"""
        rel_result = results.get('relative', {}).get('result', {})
        if method == 'pe':
            return rel_result.get('pe_ratio', 0) > 0
        elif method == 'ps':
            return rel_result.get('ps_ratio', 0) > 0
        elif method == 'pb':
            return rel_result.get('pb_ratio', 0) > 0
        elif method == 'ev':
            return rel_result.get('ev_ebitda', 0) > 0
        return False

    def get_history(self, limit: int = 50) -> List[ValuationHistoryResponse]:
        """获取历史记录列表"""
        session = Session(self.get_engine())

        try:
            records = session.query(ValuationHistory)\
                .order_by(ValuationHistory.created_at.desc())\
                .limit(limit)\
                .all()

            return [
                ValuationHistoryResponse(
                    id=r.id,
                    company_name=r.company_name,
                    industry=r.industry,
                    stage=r.stage,
                    revenue=r.revenue,
                    dcf_value=r.dcf_value,
                    dcf_wacc=r.dcf_wacc,
                    pe_value=r.pe_value,
                    ps_value=r.ps_value,
                    pb_value=r.pb_value,
                    ev_value=r.ev_value,
                    comparables_count=r.comparables_count,
                    created_at=r.created_at.isoformat() if r.created_at else None,
                )
                for r in records
            ]
        finally:
            session.close()

    def save_analysis_history(
        self,
        analysis_type: str,
        company: dict,
        results: dict
    ) -> int:
        """
        保存分析历史记录（通用方法）

        Args:
            analysis_type: 分析类型 (scenario, sensitivity, stress_test, absolute, relative)
            company: 公司信息字典
            results: 分析结果字典

        Returns:
            历史记录ID
        """
        import json
        session = Session(self.get_engine())

        try:
            # 提取基准情景估值（通常是第一个情景）
            base_value = 0
            if analysis_type == 'scenario':
                # 情景分析：使用基准情景的估值
                for name, data in results.items():
                    if name == '基准情景' or name == '基准':
                        base_value = data.get('value', 0) / 10000  # 转为万元
                        break
            elif analysis_type == 'absolute':
                base_value = results.get('result', {}).get('value', 0) / 10000
            elif analysis_type == 'relative':
                # 相对估值使用平均值或第一个方法的结果
                result = results.get('results', {})
                if isinstance(result, dict):
                    first_key = next(iter(result), None)
                    if first_key:
                        base_value = result[first_key].get('value', 0) / 10000

            # 创建历史记录
            history = ValuationHistory(
                company_name=company.get('name', ''),
                industry=company.get('industry', ''),
                stage=company.get('stage', ''),
                revenue=float(company.get('revenue', 0)) / 10000,
                net_income=float(company.get('net_income', 0)) / 10000,
                net_assets=float(company.get('net_assets', 0)) / 10000,
                ebitda=float(company.get('ebitda', 0)) / 10000,
                growth_rate=float(company.get('growth_rate', 0)),
                operating_margin=float(company.get('operating_margin', 0)),
                beta=float(company.get('beta', 0)),
                risk_free_rate=float(company.get('risk_free_rate', 0)),
                market_risk_premium=float(company.get('market_risk_premium', 0)),
                terminal_growth_rate=float(company.get('terminal_growth_rate', 0)),
                dcf_value=base_value,
                notes=json.dumps({
                    'analysis_type': analysis_type,
                    'results': results
                }, ensure_ascii=False, default=str)
            )

            session.add(history)
            session.commit()
            session.flush()
            session.refresh(history)

            return history.id
        finally:
            session.close()
