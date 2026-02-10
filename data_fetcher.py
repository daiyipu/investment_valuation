"""
Tushare数据获取模块
用于从Tushare API获取上市公司财务数据和估值倍数
"""
import tushare as ts
import pandas as pd
from typing import Optional, List, Dict, Any, Tuple
from datetime import datetime, timedelta
from models import Comparable, Company, CompanyStage


class TushareDataFetcher:
    """Tushare数据获取类"""

    # 行业分类映射
    INDUSTRY_MAP = {
        '银行': '金融',
        '保险': '金融',
        '证券': '金融',
        '房地产': '房地产',
        '汽车': '汽车',
        '医药生物': '医药',
        '计算机': '科技',
        '电子': '科技',
        '通信': '科技',
        '传媒': '传媒',
        '食品饮料': '消费',
        '家用电器': '消费',
        '纺织服装': '消费',
        '商业贸易': '消费',
        '休闲服务': '消费',
        '化工': '化工',
        '钢铁': '钢铁',
        '有色金属': '金属',
        '建筑材料': '建材',
        '建筑装饰': '建筑',
        '电气设备': '电力设备',
        '国防军工': '军工',
        '机械设备': '机械',
        '采掘': '采掘',
        '公用事业': '公用事业',
        '交通运输': '交通运输',
        '农林牧渔': '农业',
    }

    def __init__(self, token: str):
        """
        初始化Tushare API

        Args:
            token: Tushare API Token
        """
        self.pro = ts.pro_api(token)
        self.today = datetime.now().strftime('%Y%m%d')

    def get_comparable_companies(
        self,
        industry: str,
        market_cap_range: Optional[Tuple[float, float]] = None,
        limit: int = 20
    ) -> List[Comparable]:
        """
        获取同类上市公司列表

        Args:
            industry: 行业名称
            market_cap_range: 市值范围（亿元），如(10, 100)表示10-100亿
            limit: 最大返回数量

        Returns:
            可比公司列表
        """
        try:
            # 获取股票基本信息
            df = self.pro.stock_basic(
                exchange='',
                list_status='L',  # 只获取上市股票
                fields='ts_code,symbol,name,area,industry,list_date,market'
            )

            # 筛选行业
            # 简化处理：如果行业名称包含关键字
            if industry:
                df = df[df['industry'].str.contains(industry, na=False) |
                       df['industry'].isin(self._get_related_industries(industry))]

            # 获取最新交易日数据
            trade_date = self._get_latest_trade_date()

            # 获取每日行情
            df_daily = self.pro.daily_basic(
                trade_date=trade_date,
                fields='ts_code,trade_date,total_mv,circ_mv'
            )

            # 合并数据
            df = df.merge(df_daily, on='ts_code', how='inner')

            # 市值筛选
            if market_cap_range:
                min_cap, max_cap = market_cap_range
                df = df[(df['total_mv'] >= min_cap) & (df['total_mv'] <= max_cap)]

            # 排序并限制数量
            df = df.sort_values('total_mv', ascending=False).head(limit)

            # 获取财务数据
            comparables = []
            for _, row in df.iterrows():
                comp = self._get_company_financials(row['ts_code'])
                if comp:
                    comparables.append(comp)

            return comparables

        except Exception as e:
            print(f"获取可比公司失败: {e}")
            return []

    def get_financial_metrics(self, ts_code: str) -> Optional[Dict[str, Any]]:
        """
        获取单个公司的财务指标和估值倍数

        Args:
            ts_code: Tushare股票代码（如 '000001.SZ'）

        Returns:
            包含财务指标和估值倍数的字典
        """
        try:
            # 获取最新财务指标
            df_fina = self.pro.fina_indicator(
                ts_code=ts_code,
                limit=1,
                fields='ts_code,ann_date,end_date,roe,roe_waa,roe_dt,'
                       'roe_yearly,npta,npta_yearly,roe_avg,roe_yearly_avg'
            )

            # 获取最新业绩数据
            df_performance = self.pro.income(
                ts_code=ts_code,
                limit=1,
                fields='ts_code,ann_date,revenue,operate_profit,total_profit,n_income'
            )

            # 获取估值指标（日频）
            trade_date = self._get_latest_trade_date()
            df_basic = self.pro.daily_basic(
                ts_code=ts_code,
                trade_date=trade_date,
                fields='ts_code,trade_date,pe_ttm,pe,ps_ttm,ps,pb,'
                       'total_mv,circ_mv,total_share,float_share'
            )

            result = {}

            if not df_performance.empty:
                perf = df_performance.iloc[0]
                result['revenue'] = perf['revenue'] if pd.notna(perf['revenue']) else 0
                result['net_income'] = perf['n_income'] if pd.notna(perf['n_income']) else 0

            if not df_basic.empty:
                basic = df_basic.iloc[0]
                result['pe_ratio'] = basic['pe_ttm'] if pd.notna(basic['pe_ttm']) else None
                result['ps_ratio'] = basic['ps_ttm'] if pd.notna(basic['ps_ttm']) else None
                result['pb_ratio'] = basic['pb'] if pd.notna(basic['pb']) else None
                result['market_cap'] = basic['total_mv'] if pd.notna(basic['total_mv']) else None

            return result if result else None

        except Exception as e:
            print(f"获取{ts_code}财务指标失败: {e}")
            return None

    def get_industry_multiples(
        self,
        industry: str,
        method: str = 'median'
    ) -> Optional[Dict[str, float]]:
        """
        获取行业平均估值倍数

        Args:
            industry: 行业名称
            method: 统计方法，'mean' 或 'median'

        Returns:
            行业平均估值倍数字典
        """
        companies = self.get_comparable_companies(industry, limit=50)

        if not companies:
            return None

        # 收集各项倍数
        pes = [c.pe_ratio for c in companies if c.pe_ratio and c.pe_ratio > 0]
        pss = [c.ps_ratio for c in companies if c.ps_ratio and c.ps_ratio > 0]
        pbs = [c.pb_ratio for c in companies if c.pb_ratio and c.pb_ratio > 0]
        evs = [c.ev_ebitda for c in companies if c.ev_ebitda and c.ev_ebitda > 0]

        import numpy as np

        result = {}
        if pes:
            result['pe_ratio'] = np.mean(pes) if method == 'mean' else np.median(pes)
        if pss:
            result['ps_ratio'] = np.mean(pss) if method == 'mean' else np.median(pss)
        if pbs:
            result['pb_ratio'] = np.mean(pbs) if method == 'mean' else np.median(pbs)
        if evs:
            result['ev_ebitda'] = np.mean(evs) if method == 'mean' else np.median(evs)

        # 增长率
        growths = [c.growth_rate for c in companies if c.growth_rate and c.growth_rate > 0]
        if growths:
            result['avg_growth_rate'] = np.mean(growths) if method == 'mean' else np.median(growths)

        return result if result else None

    def search_by_keywords(
        self,
        keywords: List[str],
        limit: int = 10
    ) -> List[Comparable]:
        """
        按关键词搜索相关公司

        Args:
            keywords: 关键词列表
            limit: 最大返回数量

        Returns:
            匹配的公司列表
        """
        try:
            df = self.pro.stock_basic(
                exchange='',
                list_status='L',
                fields='ts_code,symbol,name,area,industry'
            )

            # 筛选包含关键词的公司
            mask = df['name'].str.contains('|'.join(keywords), na=False)
            df = df[mask].head(limit)

            comparables = []
            for _, row in df.iterrows():
                comp = self._get_company_financials(row['ts_code'])
                if comp:
                    comparables.append(comp)

            return comparables

        except Exception as e:
            print(f"关键词搜索失败: {e}")
            return []

    def get_market_data(self, ts_code: str, start_date: str, end_date: str) -> pd.DataFrame:
        """
        获取历史行情数据

        Args:
            ts_code: 股票代码
            start_date: 开始日期 YYYYMMDD
            end_date: 结束日期 YYYYMMDD

        Returns:
            行情数据DataFrame
        """
        try:
            df = self.pro.daily(
                ts_code=ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,trade_date,open,high,low,close,pre_close,'
                       'vol,amount,pct_chg'
            )
            return df.sort_values('trade_date')
        except Exception as e:
            print(f"获取行情数据失败: {e}")
            return pd.DataFrame()

    # ===== 私有辅助方法 =====

    def _get_latest_trade_date(self) -> str:
        """获取最近交易日"""
        try:
            df = self.pro.trade_cal(exchange='SSE', is_open=1)
            latest = df.iloc[-1]['cal_date']
            return latest
        except:
            # 默认返回今天
            return self.today

    def _get_related_industries(self, industry: str) -> List[str]:
        """获取相关行业列表"""
        related = []
        for key, value in self.INDUSTRY_MAP.items():
            if value == industry:
                related.append(key)
        return related

    def _get_company_financials(self, ts_code: str) -> Optional[Comparable]:
        """获取单个公司的完整财务数据"""
        try:
            # 获取基本信息
            df_basic = self.pro.stock_basic(
                ts_code=ts_code,
                fields='ts_code,symbol,name,area,industry,list_date,market'
            )

            if df_basic.empty:
                return None

            basic_info = df_basic.iloc[0]

            # 获取最新财务数据
            trade_date = self._get_latest_trade_date()

            # 日线行情
            df_daily = self.pro.daily_basic(
                ts_code=ts_code,
                trade_date=trade_date,
                fields='ts_code,pe_ttm,ps_ttm,pb,total_mv'
            )

            # 业绩数据
            df_perf = self.pro.income(
                ts_code=ts_code,
                limit=1,
                fields='ts_code,revenue,n_income,operate_profit'
            )

            # 资产负债数据
            df_bal = self.pro.balancesheet(
                ts_code=ts_code,
                limit=1,
                fields='ts_code,total_assets,total_hldr_eqy_exc_min_int'
            )

            if df_perf.empty or df_bal.empty:
                return None

            perf = df_perf.iloc[0]
            bal = df_bal.iloc[0]

            # 构建Comparable对象
            comp = Comparable(
                name=basic_info['name'],
                ts_code=ts_code,
                industry=basic_info['industry'],
                revenue=float(perf['revenue']) if pd.notna(perf['revenue']) else 0,
                net_income=float(perf['n_income']) if pd.notna(perf['n_income']) else 0,
                net_assets=float(bal['total_hldr_eqy_exc_min_int']) if pd.notna(bal['total_hldr_eqy_exc_min_int']) else 0,
            )

            # 估值倍数
            if not df_daily.empty:
                daily = df_daily.iloc[0]
                comp.pe_ratio = float(daily['pe_ttm']) if pd.notna(daily['pe_ttm']) else None
                comp.ps_ratio = float(daily['ps_ttm']) if pd.notna(daily['ps_ttm']) else None
                comp.pb_ratio = float(daily['pb']) if pd.notna(daily['pb']) else None
                comp.market_cap = float(daily['total_mv']) if pd.notna(daily['total_mv']) else None

            return comp

        except Exception as e:
            print(f"获取{ts_code}财务数据失败: {e}")
            return None


# ===== 使用示例 =====

if __name__ == "__main__":
    # 需要替换为实际的Tushare Token
    # TOKEN = "your_tushare_token_here"
    # fetcher = TushareDataFetcher(TOKEN)

    # 示例：获取软件行业可比公司
    # comparables = fetcher.get_comparable_companies("计算机", market_cap_range=(50, 500))
    # for comp in comparables:
    #     print(f"{comp.name}: P/E={comp.pe_ratio:.2f}, P/S={comp.ps_ratio:.2f}")

    # 示例：获取行业平均倍数
    # industry_multiples = fetcher.get_industry_multiples("计算机")
    # print(f"行业平均P/E: {industry_multiples['pe_ratio']:.2f}")

    print("Tushare数据获取模块已就绪")
    print("使用前请设置有效的Tushare Token")
