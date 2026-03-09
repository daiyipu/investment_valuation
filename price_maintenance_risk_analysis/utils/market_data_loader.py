# -*- coding: utf-8 -*-
"""
市场数据加载工具

功能：
- 从07_market_data_analysis生成的JSON文件加载真实市场数据
- 从Tushare获取上市公司最新财务数据
- 提供常用财务指标计算和市场数据查询
"""

import json
import os
from typing import Dict, Optional
from datetime import datetime, timedelta
import pandas as pd


def load_market_data(stock_code: str, data_dir: str = '..') -> Optional[Dict]:
    """
    加载股票的市场数据（波动率、收益率等）

    参数:
        stock_code: 股票代码，如 '300735.SZ' 或 '300735_SZ'
        data_dir: 数据文件所在目录，默认为上级目录

    返回:
        包含市场数据的字典，如果文件不存在则返回None

    使用示例:
        >>> market_data = load_market_data('300735.SZ')
        >>> if market_data:
        ...     volatility = market_data['volatility']
        ...     drift = market_data['drift']
        ...     print(f"波动率: {volatility*100:.2f}%")
        ...     print(f"收益率: {drift*100:.2f}%")
    """
    # 统一文件名格式（将点替换为下划线）
    filename = stock_code.replace('.', '_') + '_market_data.json'
    filepath = os.path.join(data_dir, filename)

    if not os.path.exists(filepath):
        print(f"⚠️ 市场数据文件不存在: {filepath}")
        print(f"   提示: 请先运行 07_market_data_analysis.ipynb 生成数据")
        return None

    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            market_data = json.load(f)

        print(f"✅ 已加载市场数据: {filepath}")
        print(f"   股票: {market_data.get('stock_name', 'N/A')} ({market_data.get('stock_code', 'N/A')})")
        print(f"   分析日期: {market_data.get('analysis_date', 'N/A')}")
        print(f"   当前价格: {market_data.get('current_price', 0):.2f} 元")

        return market_data

    except Exception as e:
        print(f"❌ 加载市场数据失败: {e}")
        return None


def get_risk_params(market_data: Dict, volatility_window: str = '60d') -> Dict:
    """
    从市场数据中提取风险参数

    参数:
        market_data: load_market_data()返回的数据
        volatility_window: 波动率窗口 ('30d', '60d', '120d', '180d')

    返回:
        风险参数字典
    """
    if not market_data:
        return {
            'volatility': 0.35,  # 默认值
            'drift': 0.08,       # 默认值
        }

    return {
        'volatility': market_data.get(f'volatility_{volatility_window}',
                                      market_data.get('volatility', 0.35)),
        'drift': market_data.get(f'annual_return_{volatility_window}',
                                 market_data.get('drift', 0.08)),
    }


def print_market_data_summary(market_data: Dict):
    """
    打印市场数据摘要
    """
    if not market_data:
        print("⚠️ 无市场数据")
        return

    print("\n" + "="*70)
    print(f"📊 {market_data.get('stock_name', 'N/A')} 市场数据摘要")
    print("="*70)

    print(f"\n📈 价格数据:")
    print(f"   当前价格: {market_data.get('current_price', 0):.2f} 元")
    print(f"   平均价格: {market_data.get('avg_price_all', 0):.2f} 元")
    print(f"   价格中位数: {market_data.get('median_price', 0):.2f} 元")

    print(f"\n⚠️ 波动率:")
    print(f"   30日波动率: {market_data.get('volatility_30d', 0)*100:.2f}%")
    print(f"   60日波动率: {market_data.get('volatility_60d', 0)*100:.2f}%")
    print(f"   120日波动率: {market_data.get('volatility_120d', 0)*100:.2f}%")
    print(f"   180日波动率: {market_data.get('volatility_180d', 0)*100:.2f}%")

    print(f"\n📊 收益率:")
    print(f"   30日年化收益: {market_data.get('annual_return_30d', 0)*100:.2f}%")
    print(f"   60日年化收益: {market_data.get('annual_return_60d', 0)*100:.2f}%")
    print(f"   120日年化收益: {market_data.get('annual_return_120d', 0)*100:.2f}%")
    print(f"   180日年化收益: {market_data.get('annual_return_180d', 0)*100:.2f}%")

    print(f"\n📈 移动平均线:")
    print(f"   MA30: {market_data.get('ma_30', 0):.2f} 元")
    print(f"   MA60: {market_data.get('ma_60', 0):.2f} 元")
    print(f"   MA120: {market_data.get('ma_120', 0):.2f} 元")
    print(f"   MA180: {market_data.get('ma_180', 0):.2f} 元")

    print(f"\n📊 其他统计:")
    print(f"   60日胜率: {market_data.get('win_rate_60d', 0)*100:.1f}%")
    print(f"   数据天数: {market_data.get('total_days', 0)} 天")

    print("="*70)


def load_market_indices_data(data_dir: str = '..') -> Optional[Dict]:
    """
    加载市场指数数据（波动率、收益率等）

    参数:
        data_dir: 数据文件所在目录，默认为上级目录

    返回:
        包含指数数据的字典，如果文件不存在则返回None

    使用示例:
        >>> indices_data = load_market_indices_data()
        >>> if indices_data:
        ...     hs300_data = indices_data['沪深300']
        ...     print(f"沪深300波动率: {hs300_data['volatility']*100:.2f}%")
        ...     print(f"沪深300收益率: {hs300_data['drift']*100:.2f}%")
    """
    filepath = os.path.join(data_dir, 'market_indices_scenario_data.json')

    if not os.path.exists(filepath):
        print(f"⚠️ 指数数据文件不存在: {filepath}")
        print(f"   提示: 请先运行 07_market_data_analysis.ipynb 第二部分生成数据")
        return None

    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            indices_data = json.load(f)

        print(f"✅ 已加载指数数据: {filepath}")
        print(f"   指数数量: {len(indices_data)}")

        return indices_data

    except Exception as e:
        print(f"❌ 加载指数数据失败: {e}")
        return None


def print_indices_data_summary(indices_data: Dict):
    """
    打印指数数据摘要
    """
    if not indices_data:
        print("⚠️ 无指数数据")
        return

    print("\n" + "="*70)
    print(f"📊 市场指数数据摘要")
    print("="*70)

    print(f"\n{'指数':<12} {'当前点位':<15} {'波动率(60日)':<15} {'收益率(60日)':<15}")
    print('-'*60)

    for name, data in indices_data.items():
        print(f"{name:<12} {data.get('current_level', 0):>10.2f}       "
              f"{data.get('volatility', 0)*100:>10.2f}%       "
              f"{data.get('drift', 0)*100:>10.2f}%")

    print("="*70)
    print(f"\n📌 可用指数:")
    for name in indices_data.keys():
        print(f"   - {name}")
    print(f"\n💡 使用示例:")
    print(f"   indices_data = load_market_indices_data()")
    print(f"   volatility = indices_data['沪深300']['volatility']")
    print(f"   drift = indices_data['沪深300']['drift']")


def get_index_risk_params(indices_data: Dict, index_name: str = '沪深300',
                          volatility_window: str = '60d') -> Dict:
    """
    从指数数据中提取指定指数的风险参数

    参数:
        indices_data: load_market_indices_data()返回的数据
        index_name: 指数名称 ('沪深300', '上证指数', 等)
        volatility_window: 波动率窗口 ('30d', '60d', '120d', '180d')

    返回:
        风险参数字典
    """
    if not indices_data or index_name not in indices_data:
        print(f"⚠️ 指数 {index_name} 数据不存在")
        return {
            'volatility': 0.20,  # 指数默认波动率较低
            'drift': 0.05,
        }

    index_data = indices_data[index_name]

    return {
        'volatility': index_data.get(f'volatility_{volatility_window}',
                                     index_data.get('volatility', 0.20)),
        'drift': index_data.get(f'return_{volatility_window}',
                                index_data.get('drift', 0.05)),
        'current_level': index_data.get('current_level', 0),
    }


def save_market_indices_data(indices_results: Dict, data_dir: str = '..'):
    """
    保存市场指数分析结果到JSON文件（统一的数据导出接口）

    参数:
        indices_results: analyze_all_indices()或批量分析返回的结果字典
        data_dir: 数据保存目录

    使用示例:
        >>> from utils.market_data_loader import save_market_indices_data
        >>> # 在notebook中分析完指数后
        >>> save_market_indices_data(all_indices_results)
    """
    # 创建简化版数据结构，供情景分析使用
    scenario_indices_data = {}

    for index_name, result in indices_results.items():
        scenario_indices_data[index_name] = {
            'code': result['index_code'],
            'name': index_name,
            'current_level': result['current_close'],

            # 波动率（多个窗口）
            'volatility_30d': result['volatility']['30日']['latest'],
            'volatility_60d': result['volatility']['60日']['latest'],
            'volatility_120d': result['volatility']['120日']['latest'],
            'volatility_180d': result['volatility']['180日']['latest'],
            'volatility': result['volatility']['60日']['latest'],  # 默认60日

            # 收益率（多个窗口）
            'return_30d': result['returns']['30日']['annualized_return'],
            'return_60d': result['returns']['60日']['annualized_return'],
            'return_120d': result['returns']['120日']['annualized_return'],
            'return_180d': result['returns']['180日']['annualized_return'],
            'drift': result['returns']['60日']['annualized_return'],  # 默认60日

            # 技术指标
            'ma_30': result['price']['30日']['latest_ma'],
            'ma_60': result['price']['60日']['latest_ma'],
            'ma_120': result['price']['120日']['latest_ma'],
            'ma_180': result['price']['180日']['latest_ma'],

            # 胜率
            'win_rate_60d': result['returns']['60日']['win_rate'],
        }

    # 保存为JSON文件
    output_file = os.path.join(data_dir, 'market_indices_scenario_data.json')
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(scenario_indices_data, f, indent=2, ensure_ascii=False)

    print(f"💾 指数数据已保存: {output_file}")
    print(f"   指数数量: {len(scenario_indices_data)}")


# 使用示例
if __name__ == '__main__':
    # 测试加载光弘科技的市场数据
    data = load_market_data('300735.SZ')

    if data:
        # 打印摘要
        print_market_data_summary(data)

        # 获取风险参数
        risk_params = get_risk_params(data, volatility_window='60d')
        print(f"\n推荐风险参数:")
        print(f"  波动率: {risk_params['volatility']*100:.2f}%")
        print(f"  漂移率: {risk_params['drift']*100:.2f}%")


# ==================== Tushare财务数据获取 ====================

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
            from datetime import datetime, timedelta

            # 获取最近4个季度的数据
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=365)).strftime('%Y%m%d')

            df = self.pro.balancesheet(
                ts_code=self.ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,ann_date,f_ann_date,end_date,total_assets,total_liab,total_hldr_eqy_exc_min_int,money_cap'
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
            from datetime import datetime, timedelta

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

    def get_all_financials(self) -> dict:
        """
        获取所有财务数据的汇总

        返回:
            包含所有财务数据的字典，格式与placement_params.json一致
        """
        print(f"正在获取 {self.ts_code} 的财务数据...")

        balancesheet = self.get_latest_balancesheet()
        income = self.get_latest_income()

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
        }

        # 计算收入增长率（同比）
        try:
            from datetime import datetime, timedelta
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
