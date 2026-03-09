#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
生成最新市场数据脚本（使用当前日期回溯）
从Tushare获取最新历史数据，计算实时指标
"""

import json
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys
import os

try:
    import tushare as ts
except ImportError:
    print("❌ 请先安装 tushare: pip install tushare")
    sys.exit(1)


def fetch_latest_data(stock_code='300735.SZ', days=500):
    """
    获取最新的历史数据（使用当前日期回溯）
    """
    print(f"📡 正在获取 {stock_code} 的最新历史数据...")

    pro = ts.pro_api()

    # 计算日期范围（使用当前日期）
    end_date = datetime.now().strftime('%Y%m%d')
    start_date = (datetime.now() - timedelta(days=days*2)).strftime('%Y%m%d')

    print(f"   查询日期范围: {start_date} ~ {end_date}")

    # 获取数据（多获取一些，确保有足够的交易日）
    df = pro.daily(ts_code=stock_code, start_date=start_date, end_date=end_date)

    if df.empty:
        print(f"❌ 未获取到数据")
        return None

    # 按日期排序
    df = df.sort_values('trade_date').reset_index(drop=True)

    print(f"✅ 获取到 {len(df)} 条数据")
    print(f"   数据范围: {df['trade_date'].iloc[0]} ~ {df['trade_date'].iloc[-1]}")

    return df


def calculate_rolling_volatility(df, window, annualize=True):
    """
    计算滚动波动率
    
    参数:
        df: 包含pct_chg列的DataFrame（pct_chg为百分比形式，如2.5表示2.5%）
        window: 窗口大小（交易日）
        annualize: 是否年化
    
    返回:
        波动率序列和统计值（小数形式，如0.23表示23%）
    """
    # Tushare的pct_chg是百分比形式（如2.5表示2.5%），需要先转换为小数
    pct_decimal = df['pct_chg'] / 100.0
    
    # 计算滚动标准差
    rolling_std = pct_decimal.rolling(window=window).std()
    
    # 年化波动率 (假设一年252个交易日)
    if annualize:
        rolling_vol = rolling_std * np.sqrt(252)
    else:
        rolling_vol = rolling_std
    
    # 返回最新值和平均值
    return {
        'latest': rolling_std.iloc[-1] * np.sqrt(252),  # 最新波动率
        'mean': rolling_std.mean() * np.sqrt(252),      # 平均波动率
        'median': rolling_std.median() * np.sqrt(252),   # 中位数波动率
        'max': rolling_std.max() * np.sqrt(252),        # 最大波动率
        'min': rolling_std.min() * np.sqrt(252),        # 最小波动率
        'series': rolling_vol                           # 完整序列（用于绘图）
    }


def calculate_annual_return(df, window):
    """
    计算年化收益率（窗口期）
    """
    if len(df) < window:
        return np.nan

    # 获取window天前的收盘价和最新收盘价
    start_price = df['close'].iloc[-window]
    end_price = df['close'].iloc[-1]

    # 计算期间收益率
    period_return = (end_price - start_price) / start_price

    # 年化收益率（假设一年252个交易日）
    years = window / 252.0
    if years > 0:
        annual_return = (1 + period_return) ** (1 / years) - 1
    else:
        annual_return = 0

    return annual_return


def generate_market_data(stock_code='300735.SZ', stock_name='光弘科技'):
    """
    生成市场数据
    """
    # 获取历史数据
    df = fetch_latest_data(stock_code)
    if df is None:
        return None

    # 基本信息
    latest_date = df['trade_date'].iloc[-1]
    current_price = df['close'].iloc[-1]

    print(f"\n📊 计算市场指标...")
    print(f"   最新交易日: {latest_date}")
    print(f"   最新收盘价: {current_price:.2f} 元")

    # 计算波动率（多窗口）
    vol_30 = calculate_rolling_volatility(df, 30)
    vol_60 = calculate_rolling_volatility(df, 60)
    vol_120 = calculate_rolling_volatility(df, 120)
    vol_180 = calculate_rolling_volatility(df, 180)

    volatility_30d = vol_30['latest']
    volatility_60d = vol_60['latest']
    volatility_120d = vol_120['latest']
    volatility_180d = vol_180['latest']

    # 计算年化收益率（多窗口）
    annual_return_30d = calculate_annual_return(df, 30)
    annual_return_60d = calculate_annual_return(df, 60)
    annual_return_120d = calculate_annual_return(df, 120)
    annual_return_180d = calculate_annual_return(df, 180)

    # 计算移动平均线
    ma_30 = df['close'].rolling(window=30).mean().iloc[-1]
    ma_60 = df['close'].rolling(window=60).mean().iloc[-1]
    ma_120 = df['close'].rolling(window=120).mean().iloc[-1]
    ma_180 = df['close'].rolling(window=180).mean().iloc[-1]

    # 计算胜率（最近60天上涨天数占比）
    win_rate_60d = (df['pct_chg'].iloc[-60:] > 0).sum() / 60

    # 计算平均价格等统计量
    avg_price_all = df['close'].mean()
    median_price = df['close'].median()
    price_std = df['close'].std()

    # 构建市场数据字典
    market_data = {
        'stock_code': stock_code,
        'stock_name': stock_name,
        'analysis_date': latest_date,
        'current_price': round(float(current_price), 2),
        'avg_price_all': round(float(avg_price_all), 2),
        'median_price': round(float(median_price), 2),
        'price_std': round(float(price_std), 2),
        'volatility_30d': round(float(volatility_30d), 4),
        'volatility_60d': round(float(volatility_60d), 4),
        'volatility_120d': round(float(volatility_120d), 4),
        'volatility_180d': round(float(volatility_180d), 4),
        'volatility': round(float(volatility_60d), 4),  # 默认使用60日波动率
        'annual_return_30d': round(float(annual_return_30d), 4),
        'annual_return_60d': round(float(annual_return_60d), 4),
        'annual_return_120d': round(float(annual_return_120d), 4),
        'annual_return_180d': round(float(annual_return_180d), 4),
        'drift': round(float(annual_return_60d), 4),  # 默认使用60日年化收益率
        'ma_30': round(float(ma_30), 2),
        'ma_60': round(float(ma_60), 2),
        'ma_120': round(float(ma_120), 2),
        'ma_180': round(float(ma_180), 2),
        'win_rate_60d': round(float(win_rate_60d), 4),
        'total_days': len(df),
        'data_source': 'tushare_realtime',
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    return market_data


def print_market_data_summary(market_data):
    """
    打印市场数据摘要
    """
    print("\n" + "="*70)
    print("📊 市场数据摘要")
    print("="*70)

    print(f"\n📈 基本信息:")
    print(f"   股票代码: {market_data['stock_code']}")
    print(f"   股票名称: {market_data['stock_name']}")
    print(f"   分析日期: {market_data['analysis_date']}")
    print(f"   当前价格: {market_data['current_price']:.2f} 元")

    print(f"\n⚠️ 波动率:")
    print(f"   30日: {market_data['volatility_30d']*100:.2f}%")
    print(f"   60日: {market_data['volatility_60d']*100:.2f}%  ← 推荐")
    print(f"   120日: {market_data['volatility_120d']*100:.2f}%")
    print(f"   180日: {market_data['volatility_180d']*100:.2f}%")

    print(f"\n📈 年化收益率:")
    print(f"   30日: {market_data['annual_return_30d']*100:.2f}%")
    print(f"   60日: {market_data['annual_return_60d']*100:.2f}%  ← 推荐")
    print(f"   120日: {market_data['annual_return_120d']*100:.2f}%")
    print(f"   180日: {market_data['annual_return_180d']*100:.2f}%")

    print(f"\n📊 移动平均线:")
    print(f"   MA30: {market_data['ma_30']:.2f} 元")
    print(f"   MA60: {market_data['ma_60']:.2f} 元")
    print(f"   MA120: {market_data['ma_120']:.2f} 元")
    print(f"   MA180: {market_data['ma_180']:.2f} 元")

    print(f"\n📈 其他指标:")
    print(f"   平均价格: {market_data['avg_price_all']:.2f} 元")
    print(f"   中位数价格: {market_data['median_price']:.2f} 元")
    print(f"   价格标准差: {market_data['price_std']:.2f} 元")
    print(f"   60日胜率: {market_data['win_rate_60d']*100:.2f}%")

    print(f"\n⏰ 数据生成时间: {market_data['generated_at']}")
    print("="*70)


if __name__ == '__main__':
    # 生成市场数据
    stock_code = '300735.SZ'
    stock_name = '光弘科技'

    print("="*70)
    print("🚀 生成最新市场数据（使用当前日期回溯）")
    print("="*70)

    # 生成数据
    market_data = generate_market_data(stock_code, stock_name)

    if market_data:
        # 打印摘要
        print_market_data_summary(market_data)

        # 保存文件
        filename = f"{stock_code.replace('.', '_')}_market_data.json"
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(market_data, f, ensure_ascii=False, indent=2)
        print(f"\n✅ 已保存市场数据到: {filename}")

        print("\n📌 使用方法:")
        print(f"   from utils.market_data_loader import load_market_data")
        print(f"   market_data = load_market_data('{stock_code}')")


# ==================== 财务数据获取 ====================

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


def generate_financial_data_json(stock_code='300735.SZ', output_file=None):
    """
    生成财务数据JSON文件

    参数:
        stock_code: 股票代码
        output_file: 输出文件路径，默认为 data/{stock_code}_financial_data.json
    """
    print("\n" + "="*70)
    print("📊 生成财务数据JSON")
    print("="*70)

    try:
        financial = TushareFinancialData(stock_code)
        data = financial.get_all_financials()

        if data:
            if output_file is None:
                output_file = f"data/{stock_code.replace('.', '_')}_financial_data.json"

            # 确保目录存在
            os.makedirs(os.path.dirname(output_file), exist_ok=True)

            # 保存为JSON
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

            print(f"\n✅ 财务数据已保存到: {output_file}")
            return data
        else:
            print("❌ 获取财务数据失败")
            return None
    except Exception as e:
        print(f"❌ 生成财务数据JSON失败: {e}")
        return None
