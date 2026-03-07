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
