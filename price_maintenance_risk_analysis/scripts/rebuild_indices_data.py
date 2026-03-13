#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
重新计算市场指数数据，使用修复后的年化收益率公式
"""

import json
import os
import numpy as np
from datetime import datetime, timedelta

def calculate_annual_return_v2(df, window):
    """
    计算年化收益率（修复后的公式）

    年化方法：
    - 60日（季度）：期间收益率 × 4
    - 120日（半年度）：期间收益率 × 2
    - 250日（年度）：期间收益率 × 1
    """
    if len(df) < window:
        return np.nan, np.nan

    # 获取window天前的收盘价和最新收盘价
    start_price = df['close'].iloc[-window]
    end_price = df['close'].iloc[-1]

    # 计算期间收益率
    period_return = (end_price - start_price) / start_price

    # 根据窗口期计算年化倍数
    if window == 60:
        # 季度：一年4个季度
        annual_return = period_return * 4
        period_return_for_table = period_return
    elif window == 120:
        # 半年度：一年2个半年度
        annual_return = period_return * 2
        period_return_for_table = period_return
    elif window == 250:
        # 年度：直接使用期间收益率
        annual_return = period_return
        period_return_for_table = period_return
    else:
        # 其他窗口：按交易日比例年化（252个交易日/年）
        years = window / 252.0
        if years > 0:
            annual_return = period_return / years
        else:
            annual_return = 0
        period_return_for_table = period_return

    return annual_return, period_return_for_table


def calculate_volatility(df, window):
    """计算波动率（年化）"""
    if len(df) < window:
        return np.nan

    # 计算对数收益率
    close_prices = df['close'].iloc[-window:].values
    log_returns = np.diff(np.log(close_prices))

    # 年化波动率（252个交易日）
    volatility = log_returns.std() * np.sqrt(252)

    return volatility


def fetch_index_data(ts_code, days=500):
    """
    获取指数历史数据

    参数:
        ts_code: 指数代码（如 '000001.SH'）
        days: 获取天数

    返回:
        DataFrame或None
    """
    try:
        import tushare as ts

        ts_token = os.environ.get('TUSHARE_TOKEN', '')
        if not ts_token:
            print("⚠️ 未设置TUSHARE_TOKEN环境变量")
            return None

        pro = ts.pro_api(ts_token)

        # 计算日期范围
        end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.now() - timedelta(days=days)).strftime('%Y%m%d')

        # 获取指数日线数据
        df = pro.index_daily(ts_code=ts_code, start_date=start_date, end_date=end_date)

        if df is None or df.empty:
            return None

        # 按日期排序
        df = df.sort_values('trade_date')
        df.reset_index(drop=True, inplace=True)

        return df

    except Exception as e:
        print(f"❌ 获取{ts_code}数据失败: {e}")
        return None


def rebuild_indices_data():
    """重新计算所有指数数据"""

    # 指数列表
    indices = [
        {'name': '上证指数', 'code': '000001.SH'},
        {'name': '深证成指', 'code': '399001.SZ'},
        {'name': '沪深300', 'code': '000300.SH'},
        {'name': '中证500', 'code': '000905.SH'},
        {'name': '创业板指', 'code': '399006.SZ'},
        {'name': '科创50', 'code': '000688.SH'},
        {'name': '中证1000', 'code': '000985.SH'},
    ]

    print("="*70)
    print("🔄 重新计算市场指数数据（使用修复后的年化收益率公式）")
    print("="*70)
    print()

    results = {}

    for index_info in indices:
        name = index_info['name']
        code = index_info['code']

        print(f"📊 获取{name}（{code}）数据...")

        # 获取历史数据
        df = fetch_index_data(code, days=500)

        if df is None or len(df) < 250:
            print(f"  ⚠️ 数据不足，跳过")
            continue

        # 计算各窗口指标
        try:
            # 当前收盘价
            current_close = df['close'].iloc[-1]

            # 波动率
            vol_20d = calculate_volatility(df, 20)
            vol_60d = calculate_volatility(df, 60)
            vol_120d = calculate_volatility(df, 120)
            vol_250d = calculate_volatility(df, 250)

            # 年化收益率和区间收益率
            ann_20d, period_20d = calculate_annual_return_v2(df, 20)
            ann_60d, period_60d = calculate_annual_return_v2(df, 60)
            ann_120d, period_120d = calculate_annual_return_v2(df, 120)
            ann_250d, period_250d = calculate_annual_return_v2(df, 250)

            # 计算MA
            ma_20 = df['close'].iloc[-20:].mean()
            ma_60 = df['close'].iloc[-60:].mean()
            ma_120 = df['close'].iloc[-120:].mean()
            ma_250 = df['close'].iloc[-250:].mean()

            # 计算胜率（60日）
            recent_60 = df['close'].iloc[-60:]
            win_days = (recent_60.diff() > 0).sum()
            win_rate_60d = win_days / 59 if len(recent_60) > 1 else 0.5

            # 构建结果
            results[name] = {
                'current_level': current_close,
                'volatility_20d': vol_20d if not np.isnan(vol_20d) else 0.30,
                'volatility_60d': vol_60d if not np.isnan(vol_60d) else 0.30,
                'volatility_120d': vol_120d if not np.isnan(vol_120d) else 0.35,
                'volatility_250d': vol_250d if not np.isnan(vol_250d) else 0.35,
                'return_20d': ann_20d if not np.isnan(ann_20d) else 0,
                'return_60d': ann_60d if not np.isnan(ann_60d) else 0,
                'return_120d': ann_120d if not np.isnan(ann_120d) else 0,
                'return_250d': ann_250d if not np.isnan(ann_250d) else 0,
                'ma_20': ma_20,
                'ma_60': ma_60,
                'ma_120': ma_120,
                'ma_250': ma_250,
                'win_rate_60d': win_rate_60d,
            }

            print(f"  ✅ 完成: 年化收益率(250日)={ann_250d*100:+.2f}%")

        except Exception as e:
            print(f"  ❌ 计算失败: {e}")
            continue

    # 保存结果
    if results:
        output_file = 'market_indices_scenario_data_v2.json'
        save_path = os.path.join('..', 'data', output_file)

        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print()
        print("="*70)
        print(f"✅ 已保存 {len(results)} 个指数数据到: {save_path}")
        print()
        print("📊 重新计算后的数据（250日年化收益率）：")
        print(f"{'指数':<12} {'当前点位':<12} {'波动率':<12} {'年化收益率':<12}")
        print('-'*70)
        for name, data in results.items():
            print(f"{name:<12} {data['current_level']:>10.2f}    {data['volatility_250d']*100:>10.2f}%    {data['return_250d']*100:>+10.2f}%")

        print()
        print("💡 新旧数据对比：")
        print("   旧方法可能使用复利公式，导致收益率偏高（如创业板90%+）")
        print("   新方法对250日直接使用期间收益率，结果更保守合理")
        print("="*70)
    else:
        print("❌ 没有成功计算任何指数数据")


if __name__ == '__main__':
    rebuild_indices_data()
