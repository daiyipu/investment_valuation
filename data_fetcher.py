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
        # 缓存最新交易日，避免重复查询
        self._cached_trade_date = None

    def get_comparable_companies(
        self,
        industry: str,
        market_cap_range: Optional[Tuple[float, float]] = None,
        limit: int = 20
    ) -> List[Comparable]:
        """
        获取同类上市公司列表

        Args:
            industry: 行业代码（如 "851041.SI"）或行业名称（如 "计算机"）
            market_cap_range: 市值范围（亿元），如(10, 100)表示10-100亿
            limit: 最大返回数量

        Returns:
            可比公司列表
        """
        try:
            # 判断是申万行业代码还是行业名称
            is_index_code = '.' in industry and industry.endswith('.SI')

            if is_index_code:
                # 使用 index_member API 获取申万行业成分股
                df_index = self.pro.index_member(
                    index_code=industry,
                    fields='con_code,index_code,in_date,out_date,is_new'
                )
                print(f"申万行业 {industry} 成分股数量: {len(df_index)}")

                if df_index.empty:
                    print(f"警告: 申万行业 {industry} 没有找到成分股")
                    return []

                # 筛选当前成分股（out_date 为空或在之后）
                df_index = df_index[(df_index['out_date'].isna()) | (df_index['out_date'] > self.today)]
                ts_codes = df_index['con_code'].tolist()
            else:
                # 使用 stock_basic API 按行业名称筛选（旧方式）
                df_basic = self.pro.stock_basic(
                    exchange='',
                    list_status='L',
                    fields='ts_code,symbol,name,area,industry,list_date,market'
                )

                # 筛选行业
                if industry:
                    df_basic = df_basic[df_basic['industry'] == industry]

                ts_codes = df_basic['ts_code'].tolist()

            print(f"筛选后的股票数量: {len(ts_codes)}")

            # 获取最新交易日数据（用于获取市值和估值倍数）
            trade_date = self._get_latest_trade_date()

            # 获取每日行情（用于市值筛选和排序）
            df_daily = self.pro.daily_basic(
                trade_date=trade_date,
                fields='ts_code,trade_date,total_mv,circ_mv'
            )

            # 筛选出目标行业的股票
            df_daily = df_daily[df_daily['ts_code'].isin(ts_codes)]

            if df_daily.empty:
                print(f"警告: 没有找到 {trade_date} 的行情数据")
                return []

            # 市值筛选
            if market_cap_range:
                min_cap, max_cap = market_cap_range
                df_daily = df_daily[(df_daily['total_mv'] >= min_cap) & (df_daily['total_mv'] <= max_cap)]

            # 按市值排序并限制数量
            df_daily = df_daily.sort_values('total_mv', ascending=False).head(limit)

            print(f"最终获取 {len(df_daily)} 家公司")

            # 获取财务数据
            comparables = []
            for ts_code in df_daily['ts_code']:
                comp = self._get_company_financials(ts_code)
                if comp:
                    comparables.append(comp)

            return comparables

        except Exception as e:
            print(f"获取可比公司失败: {e}")
            import traceback
            traceback.print_exc()
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
            print(f"\n开始查询股票{ts_code}的财务数据...")

            # 获取最新财务指标
            df_fina = self.pro.fina_indicator(
                ts_code=ts_code,
                limit=1,
                fields='ts_code,ann_date,end_date,roe,roe_waa,roe_dt,'
                       'roe_yearly,npta,npta_yearly,roe_avg,roe_yearly_avg'
            )
            print(f"  fina_indicator查询结果: {len(df_fina)}条记录" if not df_fina.empty else "  无数据")

            # 获取最新业绩数据（利润表）
            df_performance = self.pro.income(
                ts_code=ts_code,
                limit=1,
                fields='ts_code,ann_date,revenue,operate_profit,total_profit,n_income,ebitda,income_tax,int_exp,fin_exp'
            )
            print(f"  income查询结果: {len(df_performance)}条记录" if not df_performance.empty else "  无数据")


            # 获取资产负债表数据
            df_balance = self.pro.balancesheet(
                ts_code=ts_code,
                limit=1,
                fields='ts_code,ann_date,total_assets,total_hldr_eqy_exc_min_int,'
                       'total_liab,total_ncl,total_hldr_eqy_min_int,total_cur_assets'
            )

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
                # EBITDA处理：如果Tushare直接提供了ebitda，使用它；否则尝试计算
                if pd.notna(perf.get('ebitda')):
                    result['ebitda'] = perf['ebitda']
                else:
                    # 尝试计算EBITDA = 净利润 + 所得税 + 利息支出 + 财务费用
                    n_income = perf.get('n_income', 0) if pd.notna(perf.get('n_income')) else 0
                    income_tax = perf.get('income_tax', 0) if pd.notna(perf.get('income_tax')) else 0
                    int_exp = perf.get('int_exp', 0) if pd.notna(perf.get('int_exp')) else 0
                    fin_exp = perf.get('fin_exp', 0) if pd.notna(perf.get('fin_exp')) else 0

                    # 计算EBITDA
                    calculated_ebitda = n_income + income_tax + int_exp + fin_exp
                    if calculated_ebitda > 0:
                        result['ebitda'] = calculated_ebitda
                        print(f"  EBITDA已计算: {calculated_ebitda} (净利润{n_income} + 所得税{income_tax} + 利息{int_exp} + 财务费用{fin_exp})")
                    else:
                        result['ebitda'] = None
                        print(f"  无法计算EBITDA：净利润{n_income}, 所得税{income_tax}, 利息{int_exp}, 财务费用{fin_exp}")
                print(f"  最终EBITDA值: {result.get('ebitda', 'None')}")

            if not df_balance.empty:
                bal = df_balance.iloc[0]
                # 净资产 = 股东权益合计 - 少数股东权益
                result['net_assets'] = (
                    bal['total_hldr_eqy_exc_min_int'] if pd.notna(bal['total_hldr_eqy_exc_min_int']) else
                    (bal['total_hldr_eqy_min_int'] if pd.notna(bal['total_hldr_eqy_min_int']) else 0)
                )
                # 总债务 = 总负债
                result['total_debt'] = bal['total_liab'] if pd.notna(bal['total_liab']) else 0
                # 货币资金（暂时从流动资产估算，或设为0）
                result['cash_and_equivalents'] = bal.get('total_cur_assets', 0) if pd.notna(bal.get('total_cur_assets')) else 0

            if not df_basic.empty:
                basic = df_basic.iloc[0]
                result['pe_ratio'] = basic['pe_ttm'] if pd.notna(basic['pe_ttm']) else None
                result['ps_ratio'] = basic['ps_ttm'] if pd.notna(basic['ps_ttm']) else None
                result['pb_ratio'] = basic['pb'] if pd.notna(basic['pb']) else None
                result['market_cap'] = basic['total_mv'] if pd.notna(basic['total_mv']) else None

            # 获取公司基本信息（名称、行业）
            df_stock_basic = self.pro.stock_basic(
                ts_code=ts_code,
                fields='ts_code,symbol,name,industry,list_status'
            )

            if not df_stock_basic.empty:
                stock_info = df_stock_basic.iloc[0]
                result['name'] = stock_info['name']
                result['industry'] = stock_info.get('industry', '')

                # 检查股票上市状态
                list_status = stock_info.get('list_status', '')
                if list_status == 'D' or list_status == 'L':  # 退市或暂停上市
                    print(f"股票{ts_code}({stock_info['name']})上市状态: {list_status}，可能无数据")
                elif not result and not df_performance.empty:
                    print(f"股票{ts_code}({stock_info['name']})存在但财务数据为空，可能数据未更新")

            # 如果没有任何财务数据，返回None
            if not result:
                print(f"股票{ts_code}的财务数据查询结果为空")
                return None

            return result

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
        """
        获取最近的交易日（用于获取行情数据）
        使用Tushare交易日历API查询真实的最近交易日
        """
        try:
            # 如果已有缓存，直接返回
            if self._cached_trade_date:
                return self._cached_trade_date

            # 尝试使用Tushare交易日历API
            # 先获取最近10天的交易日历
            from datetime import timedelta
            end_date = datetime.now()
            start_date = end_date - timedelta(days=10)
            start_date_str = start_date.strftime('%Y%m%d')
            end_date_str = end_date.strftime('%Y%m%d')

            df_cal = self.pro.trade_cal(
                exchange='SSE',
                start_date=start_date_str,
                end_date=end_date_str,
                fields='cal_date,is_open'
            )

            if not df_cal.empty:
                # 筛选出交易日，取最近的一个
                trade_days = df_cal[df_cal['is_open'] == 1]['cal_date'].tolist()
                if trade_days:
                    latest_trade_day = trade_days[-1]
                    print(f"获取到最近交易日: {latest_trade_day}")
                    # 缓存结果
                    self._cached_trade_date = latest_trade_day
                    return latest_trade_day

            # 如果API调用失败，回退到简单逻辑
            print("警告: 交易日历API查询失败，使用简单逻辑")
            yesterday = datetime.now() - timedelta(days=1)
            yesterday_str = yesterday.strftime('%Y%m%d')
            self._cached_trade_date = yesterday_str
            return yesterday_str

        except Exception as e:
            print(f"获取交易日失败: {e}")
            # 默认返回昨天
            from datetime import timedelta
            yesterday = datetime.now() - timedelta(days=1)
            return yesterday.strftime('%Y%m%d')

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
