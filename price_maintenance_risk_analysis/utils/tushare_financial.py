# -*- coding: utf-8 -*-
"""
Tushare财务数据获取工具

功能：
- 从Tushare获取上市公司最新财务数据
- 支持资产负债表、利润表、现金流量表数据
- 提供常用财务指标计算

使用示例：
    >>> from utils.tushare_financial import TushareFinancialData
    >>> financial = TushareFinancialData('300735.SZ')
    >>> data = financial.get_latest_financials()
    >>> print(f"货币资金: {data['cash_equivalents']:.2f} 元")
"""

import os
from datetime import datetime, timedelta
import pandas as pd


class TushareFinancialData:
    """Tushare财务数据获取类"""

    def __init__(self, ts_code, token=None):
        """
        初始化

        参数:
            ts_code: 股票代码，如 '300735.SZ'
            token: Tushare Token，如果不提供则从环境变量读取
        """
        self.ts_code = ts_code
        self.token = token or os.environ.get('TUSHARE_TOKEN', '')

        if not self.token:
            raise ValueError("TUSHARE_TOKEN未设置，请设置环境变量或传入token参数")

        # 导入tushare（延迟导入）
        try:
            import tushare as ts
            ts.set_token(self.token)
            self.pro = ts.pro_api()
        except ImportError:
            raise ImportError("请安装tushare: pip install tushare")

    def get_latest_balancesheet(self) -> dict:
        """
        获取最新资产负债表数据

        返回:
            包含资产负债表主要字段的字典
        """
        try:
            # 获取最近4个季度的数据
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=365)).strftime('%Y%m%d')

            df = self.pro.balancesheet(
                ts_code=self.ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,ann_date,f_ann_date,end_date,total_assets,total_liab,total_hldr_eqy_exc_min_int'
            )

            if df.empty:
                print(f"⚠️ 未获取到{self.ts_code}的资产负债表数据")
                return None

            # 按报告期排序，取最新一期
            df = df.sort_values('end_date', ascending=False)
            latest = df.iloc[0]

            return {
                'total_assets': float(latest['total_assets']) if pd.notna(latest['total_assets']) else 0,  # 总资产
                'total_liabilities': float(latest['total_liab']) if pd.notna(latest['total_liab']) else 0,  # 总负债
                'total_equity': float(latest['total_hldr_eqy_exc_min_int']) if pd.notna(latest['total_hldr_eqy_exc_min_int']) else 0,  # 股东权益
                'cash_equivalents': float(latest['money_cap']) if 'money_cap' in latest and pd.notna(latest['money_cap']) else 0,  # 货币资金
                'report_date': latest['end_date'],
            }
        except Exception as e:
            print(f"❌ 获取资产负债表失败: {e}")
            return None

    def get_latest_income(self) -> dict:
        """
        获取最新利润表数据

        返回:
            包含利润表主要字段的字典
        """
        try:
            # 获取最近4个季度的数据
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=365)).strftime('%Y%m%d')

            df = self.pro.income(
                ts_code=self.ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,ann_date,f_ann_date,end_date,revenue,operate_profit,total_profit,n_income'
            )

            if df.empty:
                print(f"⚠️ 未获取到{self.ts_code}的利润表数据")
                return None

            # 按报告期排序，取最新一期
            df = df.sort_values('end_date', ascending=False)
            latest = df.iloc[0]

            return {
                'revenue': float(latest['revenue']) if pd.notna(latest['revenue']) else 0,  # 营业收入
                'operate_profit': float(latest['operate_profit']) if pd.notna(latest['operate_profit']) else 0,  # 营业利润
                'total_profit': float(latest['total_profit']) if pd.notna(latest['total_profit']) else 0,  # 利润总额
                'net_income': float(latest['n_income']) if pd.notna(latest['n_income']) else 0,  # 净利润
                'report_date': latest['end_date'],
            }
        except Exception as e:
            print(f"❌ 获取利润表失败: {e}")
            return None

    def get_latest_cashflow(self) -> dict:
        """
        获取最新现金流量表数据

        返回:
            包含现金流量表主要字段的字典
        """
        try:
            # 获取最近4个季度的数据
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=365)).strftime('%Y%m%d')

            df = self.pro.cashflow(
                ts_code=self.ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,ann_date,f_ann_date,end_date,n_cashflow_act'
            )

            if df.empty:
                print(f"⚠️ 未获取到{self.ts_code}的现金流量表数据")
                return None

            # 按报告期排序，取最新一期
            df = df.sort_values('end_date', ascending=False)
            latest = df.iloc[0]

            return {
                'operating_cashflow': float(latest['n_cashflow_act']) if pd.notna(latest['n_cashflow_act']) else 0,  # 经营活动现金流
                'report_date': latest['end_date'],
            }
        except Exception as e:
            print(f"❌ 获取现金流量表失败: {e}")
            return None

    def get_latest_basic_info(self) -> dict:
        """
        获取最新基本信息（市值、股本等）

        返回:
            包含基本信息的字典
        """
        try:
            df = self.pro.daily_basic(
                ts_code=self.ts_code,
                fields='ts_code,trade_date,total_mv,circ_mv,total_share,float_share,turnover_ratio'
            )

            if df.empty:
                print(f"⚠️ 未获取到{self.ts_code}的基本信息")
                return None

            # 取最新一天
            latest = df.iloc[-1]

            return {
                'total_market_cap': float(latest['total_mv']) * 10000 if pd.notna(latest['total_mv']) else 0,  # 总市值（元）
                'circulating_market_cap': float(latest['circ_mv']) * 10000 if pd.notna(latest['circ_mv']) else 0,  # 流通市值（元）
                'total_shares': float(latest['total_share']) if pd.notna(latest['total_share']) else 0,  # 总股本（万股）
                'float_shares': float(latest['float_share']) if pd.notna(latest['float_share']) else 0,  # 流通股本（万股）
                'trade_date': latest['trade_date'],
            }
        except Exception as e:
            print(f"❌ 获取基本信息失败: {e}")
            return None

    def get_all_financials(self) -> dict:
        """
        获取所有财务数据的汇总

        返回:
            包含所有财务数据的字典，格式与placement_params.json一致
        """
        print(f"正在获取 {self.ts_code} 的财务数据...")

        balancesheet = self.get_latest_balancesheet()
        income = self.get_latest_income()
        cashflow = self.get_latest_cashflow()
        basic_info = self.get_latest_basic_info()

        if not balancesheet or not income:
            print(f"❌ 获取财务数据失败")
            return None

        # 计算派生指标
        total_assets = balancesheet['total_assets']
        total_liabilities = balancesheet['total_liabilities']
        net_income = income['net_income']
        revenue = income['revenue']

        # 计算运营利润率
        operating_margin = (income['operate_profit'] / revenue) if revenue > 0 else 0

        # 计算净利润率
        net_margin = (net_income / revenue) if revenue > 0 else 0

        result = {
            # 基本信息
            'data_date': balancesheet['report_date'],

            # 资产负债表
            'total_assets': total_assets,  # 总资产（元）
            'net_assets': balancesheet['total_equity'],  # 净资产（元）
            'total_debt': total_liabilities,  # 总负债（元）
            'cash_equivalents': balancesheet['cash_equivalents'],  # 货币资金（元）
            'net_debt': total_liabilities - balancesheet['cash_equivalents'],  # 净债务 = 总负债 - 货币资金

            # 利润表
            'revenue': revenue,  # 营业收入（元）
            'operate_profit': income['operate_profit'],  # 营业利润（元）
            'net_income': net_income,  # 净利润（元）
            'operating_margin': operating_margin,  # 运营利润率
            'net_margin': net_margin,  # 净利润率

            # 现金流量表（可选）
            'operating_cashflow': cashflow['operating_cashflow'] if cashflow else 0,

            # 基本信息（股本等）
            'total_shares': int(basic_info['total_shares'] * 10000) if basic_info else 0,  # 总股本（股）
            'float_shares': int(basic_info['float_shares'] * 10000) if basic_info else 0,  # 流通股本（股）
        }

        # 计算收入增长率（同比）
        if income:
            try:
                # 获取去年同期数据
                last_year_df = self.pro.income(
                    ts_code=self.ts_code,
                    start_date=(datetime.now() - timedelta(days=400)).strftime('%Y%m%d'),
                    end_date=(datetime.now() - timedelta(days=365)).strftime('%Y%m%d'),
                    fields='ts_code,end_date,revenue'
                )
                if not last_year_df.empty:
                    last_year_revenue = last_year_df.iloc[-1]['revenue']
                    revenue_growth = (revenue - last_year_revenue) / last_year_revenue if last_year_revenue > 0 else 0
                    result['revenue_growth'] = revenue_growth
                else:
                    result['revenue_growth'] = 0.2  # 默认20%
            except:
                result['revenue_growth'] = 0.2  # 默认20%

        print(f"✅ 成功获取财务数据:")
        print(f"   总资产: {total_assets/100000000:.2f} 亿元")
        print(f"   净资产: {result['net_assets']/100000000:.2f} 亿元")
        print(f"   货币资金: {result['cash_equivalents']/100000000:.2f} 亿元")
        print(f"   净利润: {net_income/100000000:.2f} 亿元")
        print(f"   营业收入: {revenue/100000000:.2f} 亿元")

        return result


def update_placement_params_with_tushare(stock_code: str, params_file: str = None) -> dict:
    """
    使用Tushare数据更新placement_params.json文件

    参数:
        stock_code: 股票代码，如 '300735.SZ'
        params_file: params文件路径，默认为data/{stock_code}_placement_params.json

    返回:
        更新后的参数字典
    """
    import json

    if params_file is None:
        params_file = f"data/{stock_code.replace('.', '_')}_placement_params.json"

    # 检查文件是否存在
    if not os.path.exists(params_file):
        print(f"⚠️ 参数文件不存在: {params_file}")
        return None

    # 读取现有参数
    with open(params_file, 'r', encoding='utf-8') as f:
        params = json.load(f)

    # 从Tushare获取最新财务数据
    financial = TushareFinancialData(stock_code)
    tushare_data = financial.get_all_financials()

    if tushare_data:
        # 更新财务数据字段
        params.update({
            'net_assets': tushare_data['net_assets'],
            'total_debt': tushare_data['total_debt'],
            'net_income': tushare_data['net_income'],
            'cash_equivalents': tushare_data['cash_equivalents'],
            'revenue': tushare_data['revenue'],
            'operating_margin': tushare_data['operating_margin'],
            'revenue_growth': tushare_data.get('revenue_growth', params.get('revenue_growth', 0.25)),
            'data_source': 'Tushare',
            'data_date': tushare_data['data_date'],
        })

        # 保存更新后的文件
        with open(params_file, 'w', encoding='utf-8') as f:
            json.dump(params, f, indent=2, ensure_ascii=False)

        print(f"\n✅ 参数文件已更新: {params_file}")
        print(f"   数据来源: Tushare API")
        print(f"   数据日期: {tushare_data['data_date']}")

    return params


# 使用示例
if __name__ == '__main__':
    # 示例：获取光弘科技的财务数据
    stock_code = '300735.SZ'

    try:
        financial = TushareFinancialData(stock_code)
        data = financial.get_all_financials()

        if data:
            print("\n=== 财务数据汇总 ===")
            print(f"货币资金: {data['cash_equivalents']/100000000:.2f} 亿元")
            print(f"总资产: {data['total_assets']/100000000:.2f} 亿元")
            print(f"净资产: {data['net_assets']/100000000:.2f} 亿元")
            print(f"总负债: {data['total_debt']/100000000:.2f} 亿元")
            print(f"净利润: {data['net_income']/100000000:.2f} 亿元")
            print(f"营业收入: {data['revenue']/100000000:.2f} 亿元")
            print(f"运营利润率: {data['operating_margin']*100:.2f}%")
            print(f"收入增长率: {data.get('revenue_growth', 0)*100:.2f}%")
    except Exception as e:
        print(f"错误: {e}")
        print("请确保：")
        print("1. 已安装tushare: pip install tushare")
        print("2. 已设置TUSHARE_TOKEN环境变量")
