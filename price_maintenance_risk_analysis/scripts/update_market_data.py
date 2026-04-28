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
import argparse

try:
    import tushare as ts
except ImportError:
    print("❌ 请先安装 tushare: pip install tushare")
    sys.exit(1)


def _init_tushare_pro():
    """Initialize tushare pro_api with token from env or config."""
    token = os.environ.get('TUSHARE_TOKEN', '')
    if token:
        ts.set_token(token)
    return ts.pro_api()


def fetch_latest_data(stock_code='300735.SZ', days=500):
    """
    获取最新的历史数据（使用当前日期回溯）
    """
    print(f"📡 正在获取 {stock_code} 的最新历史数据...")

    pro = _init_tushare_pro()

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
    计算滚动波动率（基于连续复利，对数收益率）

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    【重要说明】波动率计算方法

    本函数使用**连续复利（对数收益率）**计算波动率，符合GBM模型要求。

    计算公式：
        对数收益率 = ln(P_t / P_{t-1})
        波动率 = std(对数收益率) × √250

    注意：
        - 波动率必须使用连续复利计算（与收益率不同）
        - 年化因子使用250（中国A股年交易日数）
        - 此波动率可直接用于GBM模型、蒙特卡洛模拟等

    理论依据：
        GBM模型中，波动率σ是对数收益率的标准差
        这是伊藤引理的要求，不能使用离散复利

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    参数:
        df: 包含close列的DataFrame（收盘价数据）
        window: 窗口大小（交易日）
        annualize: 是否年化（默认True）

    返回:
        dict: 波动率序列和统计值（小数形式，如0.23表示23%）
            - latest: 最新波动率
            - mean: 平均波动率
            - median: 中位数波动率
            - max: 最大波动率
            - min: 最小波动率
            - series: 完整序列（用于绘图）

    示例:
        >>> vol = calculate_rolling_volatility(df, 60)
        >>> print(f"60日波动率: {vol['latest']*100:.2f}%")
        60日波动率: 33.80%
        >>> # 可直接用于GBM模型
        >>> drift = ...  # 漂移率
        >>> volatility = vol['latest']  # 波动率
        >>> monte_carlo_simulation(drift=drift, volatility=volatility)
    """
    # 计算对数收益率（连续复利）
    # 注意：波动率必须使用连续复利计算
    log_returns = np.log(df['close']).diff()

    # 计算滚动标准差
    rolling_std = log_returns.rolling(window=window).std()

    # 年化波动率（年化因子250，中国A股年交易日数）
    if annualize:
        rolling_vol = rolling_std * np.sqrt(250)
    else:
        rolling_vol = rolling_std

    # 返回最新值和平均值
    return {
        'latest': rolling_std.iloc[-1] * np.sqrt(250),  # 最新波动率
        'mean': rolling_std.mean() * np.sqrt(250),      # 平均波动率
        'median': rolling_std.median() * np.sqrt(250),   # 中位数波动率
        'max': rolling_std.max() * np.sqrt(250),        # 最大波动率
        'min': rolling_std.min() * np.sqrt(250),        # 最小波动率
        'series': rolling_vol                           # 完整序列（用于绘图）
    }


def calculate_annual_return(df, window):
    """
    计算年化收益率（基于离散复利，用于报告展示）

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    【重要说明】年化收益率计算方法

    本函数计算**离散复利的年化收益率**，用于报告展示。

    计算公式（离散复利）：
        年化收益率 = 区间收益率 × (250 / 窗口期天数)

    注意：
        - 此方法保持与区间收益率的一致性，便于理解
        - 年化因子使用250（中国A股年交易日数）
        - 此方法用于报告展示，直观但不严谨

    【GBM模型使用】
    如需在GBM模型中使用年化漂移率，请使用：
        from utils.return_conversion import get_gbm_drift_from_discrete

        # 方法1：直接从离散收益率计算GBM漂移率（推荐）
        drift = get_gbm_drift_from_discrete(period_return, window)

        # 方法2：手动转换
        from utils.return_conversion import discrete_to_continuous
        period_continuous = discrete_to_continuous(period_return)
        drift = period_continuous * (250 / window)

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    参数:
        df: DataFrame，包含收盘价数据
        window: 时间窗口（交易日）

    返回:
        float: 年化离散复利收益率

    示例:
        >>> df = ...  # 包含收盘价数据
        >>> annual_20d = calculate_annual_return(df, 20)
        >>> print(f"20日年化收益率: {annual_20d*100:.2f}%")
        20日年化收益率: -219.96%

        >>> # 用于GBM模型时需要转换
        >>> period_return = calculate_period_return(df, 20)  # -17.60%
        >>> drift = get_gbm_drift_from_discrete(period_return, 20)
        >>> print(f"GBM漂移率（连续复利年化）: {drift*100:.2f}%")
        GBM漂移率（连续复利年化）: -241.88%
    """
    if len(df) < window:
        return np.nan

    # 获取window天前的收盘价和最新收盘价
    start_price = df['close'].iloc[-window]
    end_price = df['close'].iloc[-1]

    # 步骤1: 计算区间离散复利收益率
    period_discrete_return = (end_price - start_price) / start_price

    # 步骤2: 年化（离散复利）
    # 注意：这是离散复利的年化，用于报告展示
    # 年化因子使用250（中国A股年交易日数）
    annual_discrete_return = period_discrete_return * (250.0 / window)

    return annual_discrete_return


def calculate_period_return(df, window):
    """
    计算区间收益率（离散复利，用于报告展示）

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    【重要说明】离散复利 vs 连续复利

    本函数返回**离散复利收益率**（简单涨跌幅），用于报告展示，直观易懂。

    离散复利（本函数）：
        公式: (P_end - P_start) / P_start
        用途: 报告展示、用户界面
        优点: 直观易懂，符合日常理解

    连续复利（对数收益率）：
        公式: ln(P_end / P_start)
        用途: GBM模型、蒙特卡洛模拟、VaR计算
        优点: 数学性质好，时间可加，符合伊藤引理

    【转换方法】
    如需在GBM模型中使用，请转换为连续复利：
        from utils.return_conversion import discrete_to_continuous

        # 展示层：离散复利（报告中的数据）
        period_return_discrete = calculate_period_return(df, 20)  # -17.60%

        # 计算层：转换为连续复利（GBM模型）
        period_return_continuous = discrete_to_continuous(period_return_discrete)
        # = ln(1 - 0.176) = -19.35%

        # 年化连续复利（用于GBM漂移率）
        drift = period_return_continuous * (250 / 20)

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    参数:
        df: DataFrame，包含收盘价数据
        window: 时间窗口（交易日）

    返回:
        float: 区间离散复利收益率（简单涨跌幅）

    示例:
        >>> df = ...  # 包含收盘价数据
        >>> return_20d = calculate_period_return(df, 20)
        >>> print(f"20日收益率: {return_20d*100:.2f}%")
        20日收益率: -17.60%

        >>> # 用于GBM模型时需要转换
        >>> from utils.return_conversion import discrete_to_continuous, get_gbm_drift_from_discrete
        >>> drift = get_gbm_drift_from_discrete(return_20d, 20)
        >>> print(f"GBM漂移率: {drift*100:.2f}%")
        GBM漂移率: -241.88%
    """
    if len(df) < window:
        return np.nan

    # 获取window天前的收盘价和最新收盘价
    start_price = df['close'].iloc[-window]
    end_price = df['close'].iloc[-1]

    # 计算离散复利收益率（简单涨跌幅），用于报告展示
    # 注意：这是离散复利，不是连续复利（对数收益率）
    period_discrete_return = (end_price - start_price) / start_price

    return period_discrete_return


def fetch_market_turnover_data(target_days=1200):
    """
    获取市场换手率数据（上海+深圳）

    参数:
        target_days: 目标获取天数（默认1200个交易日≈5年）

    返回:
        dict: 包含当前换手率、历史数据和分位数
    """
    try:
        import tushare as ts
        pro = _init_tushare_pro()

        print(f"\n📡 正在获取市场换手率数据...")

        # 使用批量获取方式
        end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.now() - timedelta(days=365*6)).strftime('%Y%m%d')  # 回溯6年确保有足够交易日

        print(f"   查询日期范围: {start_date} ~ {end_date}")

        # 批量获取上海A股数据
        df_sh_all = pro.daily_info(start_date=start_date, end_date=end_date, exchange='SH',
                                    fields='trade_date,ts_name,total_mv,float_mv,tr')

        # 批量获取深圳市场数据
        df_sz_all = pro.daily_info(start_date=start_date, end_date=end_date, exchange='SZ',
                                    fields='trade_date,ts_name,total_mv,float_mv,amount')

        print(f"✅ 获取到上海数据 {len(df_sh_all)} 条，深圳数据 {len(df_sz_all)} 条")

        # 只使用深圳市场数据（数据历史更长，约5年）
        df_sz = df_sz_all[df_sz_all['ts_name'] == '深圳市场'].copy()

        if len(df_sz) == 0:
            print(f"⚠️ 没有找到深圳市场数据")
            return None

        # 计算深圳市场换手率（作为全市场换手率）
        df_sz['sz_turnover'] = (df_sz['amount'] / df_sz['float_mv'] * 100).fillna(3.0)

        # 使用深圳数据作为市场数据
        df_merged = df_sz[['trade_date', 'sz_turnover']].copy()
        df_merged = df_merged.rename(columns={'sz_turnover': 'weighted_turnover'})

        # 按日期排序
        df_merged = df_merged.sort_values('trade_date').reset_index(drop=True)

        # 获取最近120个交易日的中位数作为当前换手率
        recent_120_days = df_merged.iloc[-120:] if len(df_merged) >= 120 else df_merged
        market_turnover = recent_120_days['weighted_turnover'].median()
        latest_turnover = df_merged.iloc[-1]['weighted_turnover']  # 最新一天数据，用于参考

        print(f"✅ 成功获取市场换手率数据（使用深圳市场数据代表全市场）：")
        print(f"   数据量：{len(df_merged)} 个交易日")
        print(f"   数据范围：{df_merged['trade_date'].iloc[0]} ~ {df_merged['trade_date'].iloc[-1]}")
        print(f"   最新一日换手率：{latest_turnover:.2f}%")
        print(f"   当前市场换手率（120日中位数）：{market_turnover:.2f}%")

        # 获取最近target_days天的数据计算分位数
        if len(df_merged) >= target_days:
            recent_df = df_merged.iloc[-target_days:]
        else:
            recent_df = df_merged
            print(f"⚠️ 数据量不足{target_days}天，使用全部{len(df_merged)}天数据")

        historical_series = recent_df['weighted_turnover']
        # 计算120日中位数在历史数据中的分位数
        current_percentile = (historical_series < market_turnover).mean() * 100

        print(f"✅ 历史分位数计算完成（120日中位数在{len(recent_df)}天历史中的分位数）：{current_percentile:.1f}%")

        # 保存历史换手率数据（用于绘制图表）
        historical_turnover_data = []
        for _, row in recent_df.iterrows():
            historical_turnover_data.append({
                'date': row['trade_date'],
                'turnover': round(float(row['weighted_turnover']), 4)
            })

        return {
            'current_turnover': round(float(market_turnover), 4),
            'latest_turnover': round(float(latest_turnover), 4),  # 最新一日数据
            'historical_count': len(recent_df),
            'historical_percentile': round(float(current_percentile), 2),
            'historical_mean': round(float(historical_series.mean()), 4),
            'historical_median': round(float(historical_series.median()), 4),
            'historical_std': round(float(historical_series.std()), 4),
            'historical_min': round(float(historical_series.min()), 4),
            'historical_max': round(float(historical_series.max()), 4),
            'historical_data': historical_turnover_data,  # 新增：历史数据数组
            'data_fetched_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

    except Exception as e:
        print(f"⚠️ 获取市场换手率数据失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def calculate_ma20_from_base_date(df, base_date_str):
    """
    基于基准日期（发行日/报价日）计算MA20价格（成交量加权均价）

    规则：从发行日前一个交易日（含）开始往前累计20个交易日
         如果发行日当天不是交易日或发行日是未来日期，使用最新可用交易日
         使用成交量加权均价：累计成交额/累计成交量

    参数:
        df: 股票历史数据DataFrame
        base_date_str: 基准日期字符串（YYYYMMDD格式）

    返回:
        float: MA20价格（成交量加权均价），如果计算失败返回None
    """
    try:
        import pandas as pd

        # 确保数据按日期排序
        df_sorted = df.sort_values('trade_date').reset_index(drop=True)

        # 找到发行日或之前的最近交易日作为起点
        df_before_base = df_sorted[df_sorted['trade_date'] <= base_date_str]

        # 如果发行日是未来日期或没有数据，使用最新交易日
        if df_before_base.empty:
            print(f"   发行日{base_date_str}是未来日期或无数据，使用最新交易日计算MA20")
            # 使用最新交易日之前的20个交易日（不包括最新交易日）
            if len(df_sorted) >= 21:
                # 最新交易日之前20个交易日
                df_20_days = df_sorted.iloc[-21:-1]  # 排除最后一天（最新交易日）
                # 计算成交量加权均价
                total_amount = df_20_days['amount'].sum()  # 千元
                total_vol = df_20_days['vol'].sum()        # 手

                if total_vol > 0:
                    # 加权均价（元）= 累计成交额(千元) / 累计成交量(手) * 10
                    ma20_price = (total_amount / total_vol) * 10
                else:
                    # 如果成交量为0（极罕见），使用简单收盘价平均
                    ma20_price = df_20_days['close'].mean()

                actual_start_date = df_20_days['trade_date'].iloc[0]
                actual_end_date = df_20_days['trade_date'].iloc[-1]
                print(f"   计算MA20（使用最新20个交易日）...")
                print(f"   实际使用区间：{actual_start_date} ~ {actual_end_date}（最新交易日前20个交易日，不含最新交易日）")
                print(f"   计算方式：成交量加权均价")
                return float(ma20_price)
            else:
                print(f"警告：历史交易日数据不足21天，无法计算MA20")
                return None

        # 从发行日或之前的最近交易日开始，往前找20个交易日（不包括基准日当天）
        start_index = df_before_base.index[-1]

        # 检查是否有足够的交易日数据（需要基准日前20个交易日）
        if start_index >= 20:
            # 有足够的交易日数据，往前取20个交易日（不包括基准日）
            df_20_days = df_sorted.iloc[start_index - 20:start_index]

            # 计算成交量加权均价
            # vol: 成交量（手），amount: 成交额（千元）
            # 加权均价 = 累计成交额 / 累计成交量
            # 注意单位转换：amount是千元，vol是手
            # 需要将成交额转换为元，或者保持一致的单位
            total_amount = df_20_days['amount'].sum()  # 千元
            total_vol = df_20_days['vol'].sum()        # 手

            if total_vol > 0:
                # 加权均价（元）= (累计成交额(千元) * 1000) / (累计成交量(手) * 100)
                # 简化：加权均价 = 累计成交额(千元) / 累计成交量(手) * 10
                ma20_price = (total_amount / total_vol) * 10
            else:
                # 如果成交量为0（极罕见），使用简单收盘价平均
                ma20_price = df_20_days['close'].mean()

            actual_start_date = df_20_days['trade_date'].iloc[0]
            actual_end_date = df_20_days['trade_date'].iloc[-1]
            print(f"   计算MA20（基于发行日{base_date_str}）...")
            print(f"   实际使用区间：{actual_start_date} ~ {actual_end_date}（基准日前20个交易日，不含基准日）")
            print(f"   计算方式：成交量加权均价")
            return float(ma20_price)
        else:
            # 数据不足，使用最新20个交易日作为fallback（排除最新交易日）
            print(f"   从发行日{base_date_str}往前推的数据不足20天，使用最新交易日前20个交易日")
            if len(df_sorted) >= 21:
                # 最新交易日之前20个交易日
                df_20_days = df_sorted.iloc[-21:-1]  # 排除最后一天（最新交易日）
                total_amount = df_20_days['amount'].sum()  # 千元
                total_vol = df_20_days['vol'].sum()        # 手

                if total_vol > 0:
                    ma20_price = (total_amount / total_vol) * 10
                else:
                    ma20_price = df_20_days['close'].mean()

                actual_start_date = df_20_days['trade_date'].iloc[0]
                actual_end_date = df_20_days['trade_date'].iloc[-1]
                print(f"   实际使用区间：{actual_start_date} ~ {actual_end_date}（最新交易日前20个交易日，不含最新交易日）")
                print(f"   计算方式：成交量加权均价")
                return float(ma20_price)
            else:
                print(f"警告：历史交易日数据不足21天，无法计算MA20")
                return None

    except Exception as e:
        print(f"计算MA20失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def generate_market_data(stock_code='300735.SZ', stock_name='光弘科技', issue_date=None):
    """
    生成市场数据

    参数:
        stock_code: 股票代码
        stock_name: 股票名称
        issue_date: 发行日/报价日（格式YYYYMMDD），用于计算MA20
                   如果为None，则使用当前日期作为发行日
    """
    # 获取历史数据
    df = fetch_latest_data(stock_code)
    if df is None:
        return None

    # 确定发行日（报价日）
    if issue_date is None:
        issue_date = datetime.now().strftime('%Y%m%d')
        print(f"   使用当前日期作为发行日：{issue_date}")
    else:
        print(f"   使用指定发行日：{issue_date}")

    # 计算询价邀请日（发行日-3天）
    from datetime import timedelta
    issue_date_obj = datetime.strptime(issue_date, '%Y%m%d')
    invitation_date = (issue_date_obj - timedelta(days=3)).strftime('%Y%m%d')
    print(f"   询价邀请日：{invitation_date}（发行日-3天）")

    # 基本信息 - 获取指定发行日的价格
    # 首先尝试获取指定发行日的价格，如果当天不是交易日，获取最近的交易日价格
    df_sorted = df.sort_values('trade_date').reset_index(drop=True)

    # 找到发行日或之后的最近交易日
    df_on_or_after_issue = df_sorted[df_sorted['trade_date'] >= issue_date]
    if not df_on_or_after_issue.empty:
        # 有发行日或之后的数据，使用第一个交易日（最接近发行日）
        latest_date = df_on_or_after_issue['trade_date'].iloc[0]
        current_price = df_on_or_after_issue['close'].iloc[0]
        print(f"   使用{latest_date}的价格（最接近发行日{issue_date}）")
    else:
        # 没有发行日或之后的数据，使用最新交易日
        latest_date = df_sorted['trade_date'].iloc[-1]
        current_price = df_sorted['close'].iloc[-1]
        print(f"   发行日{issue_date}暂无数据，使用最新交易日{latest_date}的价格")

    print(f"\n📊 计算市场指标...")
    print(f"   发行日: {issue_date}")
    print(f"   实际使用交易日: {latest_date}")
    print(f"   实际使用收盘价: {current_price:.2f} 元")

    # 计算MA20（基于发行日）
    print(f"   计算MA20（基于发行日{issue_date}）...")
    ma_20 = calculate_ma20_from_base_date(df, issue_date)
    if ma_20 is None:
        print(f"无法基于发行日{issue_date}计算MA20")
        return None
    print(f"   MA20（基于发行日{issue_date}的成交量加权均价）: {ma_20:.2f} 元")

    # 计算波动率（多窗口）
    vol_20 = calculate_rolling_volatility(df, 20)
    vol_60 = calculate_rolling_volatility(df, 60)
    vol_120 = calculate_rolling_volatility(df, 120)
    vol_250 = calculate_rolling_volatility(df, 250)

    volatility_20d = vol_20['latest']
    volatility_60d = vol_60['latest']
    volatility_120d = vol_120['latest']
    volatility_250d = vol_250['latest']

    # 计算年化收益率（多窗口）
    annual_return_20d = calculate_annual_return(df, 20)
    annual_return_60d = calculate_annual_return(df, 60)
    annual_return_120d = calculate_annual_return(df, 120)
    annual_return_250d = calculate_annual_return(df, 250)

    # 计算区间收益率（多窗口）
    period_return_20d = calculate_period_return(df, 20)
    period_return_60d = calculate_period_return(df, 60)
    period_return_120d = calculate_period_return(df, 120)
    period_return_250d = calculate_period_return(df, 250)

    # 计算移动平均线
    ma_20 = df['close'].rolling(window=20).mean().iloc[-1]
    ma_30 = df['close'].rolling(window=30).mean().iloc[-1]
    ma_60 = df['close'].rolling(window=60).mean().iloc[-1]
    ma_120 = df['close'].rolling(window=120).mean().iloc[-1]
    ma_250 = df['close'].rolling(window=250).mean().iloc[-1]

    # 计算胜率（多个时间窗口的上涨天数占比）
    win_rate_20d = (df['pct_chg'].iloc[-20:] > 0).sum() / 20 if len(df) >= 20 else 0
    win_rate_60d = (df['pct_chg'].iloc[-60:] > 0).sum() / 60 if len(df) >= 60 else 0
    win_rate_120d = (df['pct_chg'].iloc[-120:] > 0).sum() / 120 if len(df) >= 120 else 0
    win_rate_250d = (df['pct_chg'].iloc[-250:] > 0).sum() / 250 if len(df) >= 250 else 0

    # 计算平均价格等统计量
    avg_price_all = df['close'].mean()
    median_price = df['close'].median()
    price_std = df['close'].std()

    # 保存价格序列（用于时间序列预测模型，如ARIMA/GARCH）
    # 保存最近500个交易日的收盘价
    price_series = df['close'].iloc[-500:].tolist() if len(df) >= 500 else df['close'].tolist()

    # 获取市场换手率数据
    print(f"\n📊 获取市场换手率数据...")
    turnover_data = fetch_market_turnover_data(target_days=1200)  # 获取5年数据
    if turnover_data:
        print(f"✅ 市场换手率数据获取成功")
    else:
        print(f"⚠️ 市场换手率数据获取失败，将使用默认值")

    # 构建市场数据字典
    market_data = {
        'stock_code': stock_code,
        'stock_name': stock_name,
        'analysis_date': datetime.now().strftime('%Y%m%d'),  # 使用当前日期，而不是最后交易日期
        'latest_trading_date': latest_date,  # 新增：最后交易日期字段
        'issue_date': issue_date,  # 新增：发行日/报价日
        'invitation_date': invitation_date,  # 新增：询价邀请日
        'current_price': round(float(current_price), 2),
        'avg_price_all': round(float(avg_price_all), 2),
        'median_price': round(float(median_price), 2),
        'price_std': round(float(price_std), 2),
        'volatility_20d': round(float(volatility_20d), 4),
        'volatility_60d': round(float(volatility_60d), 4),
        'volatility_120d': round(float(volatility_120d), 4),
        'volatility_250d': round(float(volatility_250d), 4),
        'volatility': round(float(volatility_60d), 4),  # 默认使用60日波动率
        'annual_return_20d': round(float(annual_return_20d), 4),
        'annual_return_60d': round(float(annual_return_60d), 4),
        'annual_return_120d': round(float(annual_return_120d), 4),
        'annual_return_250d': round(float(annual_return_250d), 4),
        'period_return_20d': round(float(period_return_20d), 4),
        'period_return_60d': round(float(period_return_60d), 4),
        'period_return_120d': round(float(period_return_120d), 4),
        'period_return_250d': round(float(period_return_250d), 4),
        'drift': round(float(annual_return_60d), 4),  # 默认使用60日年化收益率
        'ma_20': round(float(ma_20), 2),
        'ma_30': round(float(ma_30), 2),
        'ma_60': round(float(ma_60), 2),
        'ma_120': round(float(ma_120), 2),
        'ma_250': round(float(ma_250), 2),
        'win_rate_20d': round(float(win_rate_20d), 4),
        'win_rate_60d': round(float(win_rate_60d), 4),
        'win_rate_120d': round(float(win_rate_120d), 4),
        'win_rate_250d': round(float(win_rate_250d), 4),
        'total_days': len(df),
        'price_series': price_series,  # 新增：完整价格序列，用于时间序列预测
        'data_source': 'tushare_realtime',
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    # 添加市场换手率数据（如果获取成功）
    if turnover_data:
        market_data['market_turnover'] = turnover_data

    return market_data

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
    print(f"   分析日期: {market_data['analysis_date']} (数据生成日期)")
    if 'latest_trading_date' in market_data:
        print(f"   最后交易日: {market_data['latest_trading_date']} (市场数据日期)")
    print(f"   当前价格: {market_data['current_price']:.2f} 元")

    print(f"\n⚠️ 波动率:")
    print(f"   月度(20日): {market_data['volatility_20d']*100:.2f}%")
    print(f"   季度(60日): {market_data['volatility_60d']*100:.2f}%  ← 推荐")
    print(f"   半年(120日): {market_data['volatility_120d']*100:.2f}%")
    print(f"   年度(250日): {market_data['volatility_250d']*100:.2f}%")

    print(f"\n📈 年化收益率（连续复利）:")
    print(f"   月度(20日): {market_data['annual_return_20d']*100:.2f}%")
    print(f"   季度(60日): {market_data['annual_return_60d']*100:.2f}%  ← 推荐")
    print(f"   半年(120日): {market_data['annual_return_120d']*100:.2f}%")
    print(f"   年度(250日): {market_data['annual_return_250d']*100:.2f}%")
    print(f"\n   注意：年化收益率采用连续复利（对数收益率）计算，符合GBM模型要求")

    print(f"\n📊 移动平均线:")
    print(f"   MA20: {market_data['ma_20']:.2f} 元")
    print(f"   MA30: {market_data['ma_30']:.2f} 元")
    print(f"   MA60: {market_data['ma_60']:.2f} 元")
    print(f"   MA120: {market_data['ma_120']:.2f} 元")
    print(f"   MA250: {market_data['ma_250']:.2f} 元")

    print(f"\n📈 其他指标:")
    print(f"   平均价格: {market_data['avg_price_all']:.2f} 元")
    print(f"   中位数价格: {market_data['median_price']:.2f} 元")
    print(f"   价格标准差: {market_data['price_std']:.2f} 元")

    print(f"\n🎯 胜率（上涨天数占比）:")
    print(f"   月度(20日): {market_data.get('win_rate_20d', 0)*100:.1f}%")
    print(f"   季度(60日): {market_data.get('win_rate_60d', 0)*100:.1f}%")
    print(f"   半年(120日): {market_data.get('win_rate_120d', 0)*100:.1f}%")
    print(f"   年度(250日): {market_data.get('win_rate_250d', 0)*100:.1f}%")

    # 显示市场换手率数据
    if 'market_turnover' in market_data:
        turnover = market_data['market_turnover']
        current_to = turnover.get('current_turnover', 0)
        print(f"\n💹 市场换手率（二级市场活跃度）:")
        print(f"   当前换手率: {current_to:.2f}%")  # current_turnover已经是百分比形式
        if turnover.get('historical_count', 0) > 0:
            print(f"   历史分位数: {turnover.get('historical_percentile', 0):.1f}%（最近{turnover.get('historical_count', 0)}个交易日）")
            print(f"   历史均值: {turnover.get('historical_mean', 0):.2f}%")
            print(f"   历史中位数: {turnover.get('historical_median', 0):.2f}%")
        else:
            print(f"   历史数据: 不可用")

    print(f"\n⏰ 数据生成时间: {market_data['generated_at']}")
    print("="*70)


def update_placement_params_issue_price(stock_code, ma20, data_dir):
    """
    根据MA20更新placement_params.json中的issue_price

    参数:
        stock_code: 股票代码
        ma20: MA20价格
        data_dir: 数据目录
    """
    try:
        placement_file = os.path.join(data_dir, f"{stock_code.replace('.', '_')}_placement_params.json")

        if not os.path.exists(placement_file):
            print(f"⚠️ 未找到 {placement_file}")
            print(f"   路径说明: placement_params.json文件不存在")
            print(f"   建议: 运行一次报告生成脚本，会自动创建该文件")
            print(f"   跳过发行价更新")
            return

        # 读取placement_params.json
        with open(placement_file, 'r', encoding='utf-8') as f:
            placement_params = json.load(f)

        # 计算新的发行价（MA20的9折，即10%折价）
        new_issue_price = ma20 * 0.9
        old_issue_price = placement_params.get('issue_price')

        # 更新issue_price和price_source
        placement_params['issue_price'] = new_issue_price
        placement_params['price_source'] = 'MA20的9折'

        # 保存
        with open(placement_file, 'w', encoding='utf-8') as f:
            json.dump(placement_params, f, indent=2, ensure_ascii=False)

        print(f"✅ 已更新发行价:")
        print(f"   旧发行价: {old_issue_price:.2f} 元")
        print(f"   新发行价: {new_issue_price:.2f} 元（MA20: {ma20:.2f} × 0.9）")
        print(f"   折价率: -10.0%")
        print(f"   文件: {placement_file}")

    except Exception as e:
        print(f"⚠️ 更新发行价失败: {e}")


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

            # tushare的资产负债表数据单位已经是元，不需要转换
            print(f"   总资产: {float(latest['total_assets'])/100000000:.2f} 亿元")
            print(f"   总负债: {float(latest['total_liab'])/100000000:.2f} 亿元")

            return {
                'total_assets': float(latest['total_assets']) if pd.notna(latest['total_assets']) else 0,  # 总资产（元）
                'total_liabilities': float(latest['total_liab']) if pd.notna(latest['total_liab']) else 0,  # 总负债（元）
                'total_equity': float(latest['total_hldr_eqy_exc_min_int']) if pd.notna(latest['total_hldr_eqy_exc_min_int']) else 0,  # 股东权益（元）
                'cash_equivalents': float(latest['money_cap']) if 'money_cap' in latest and pd.notna(latest['money_cap']) else 0,  # 货币资金（元）
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
                fields='ts_code,ann_date,f_ann_date,end_date,revenue,operate_profit,total_profit,n_income_attr_p'
            )

            if df.empty:
                print(f"⚠️ 未获取到{self.ts_code}的利润表数据")
                return None

            # 按报告期排序，取最新一期
            df = df.sort_values('end_date', ascending=False)
            latest = df.iloc[0]

            # tushare的利润表数据单位已经是元，不需要转换
            # n_income_attr_p字段: 归属母公司所有者的净利润（不含少数股东损益）
            # 这是分析上市公司时最常用的净利润指标
            net_income_yuan = float(latest['n_income_attr_p']) if pd.notna(latest['n_income_attr_p']) else 0

            print(f"从Tushare获取利润表数据成功，期间: {latest['end_date']}")
            print(f"   字段说明：n_income_attr_p = 归属母公司所有者的净利润（不含少数股东损益）")
            print(f"   归母净利润: {net_income_yuan/100000000:.2f} 亿元")

            print(f"   原始净利润（元）: {net_income_yuan:.0f}")
            print(f"   转换后净利润（亿元）: {net_income_yuan/100000000:.2f}")

            return {
                'revenue': float(latest['revenue']) if pd.notna(latest['revenue']) else 0,  # 营业收入（元）
                'operate_profit': float(latest['operate_profit']) if pd.notna(latest['operate_profit']) else 0,  # 营业利润（元）
                'total_profit': float(latest['total_profit']) if pd.notna(latest['total_profit']) else 0,  # 利润总额（元）
                'net_income': net_income_yuan,  # 净利润（元）
                'report_date': latest['end_date'],
            }
        except Exception as e:
            print(f"❌ 获取利润表失败: {e}")
            return None

    def get_historical_net_income(self, years=5) -> list:
        """
        获取历史净利润数据，用于计算CAGR
        优先使用年度数据，如某年无年度数据则使用最新季度数据年化

        参数:
            years: 获取最近几年的数据（默认5年）

        返回:
            历史净利润列表，从最新到最旧排序
        """
        try:
            # 获取最近几年的数据（多取一些以确保有足够数据）
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=(years+2)*365)).strftime('%Y%m%d')

            print(f"正在获取历史净利润数据...")
            print(f"   字段说明：n_income_attr_p = 归属母公司所有者的净利润（不含少数股东损益）")
            print(f"   查询时间范围: {start_date} 至 {end_date}")
            print(f"   目标年数: {years}年")

            df = self.pro.income(
                ts_code=self.ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,end_date,n_income_attr_p'
            )

            if df.empty:
                print(f"⚠️ 未获取到{self.ts_code}的历史净利润数据")
                return None

            # tushare的利润表数据单位已经是元，不需要转换
            print(f"   注意：tushare数据单位已经是元，无需转换")
            sample = df.iloc[0]['n_income_attr_p'] if len(df) > 0 else 0
            print(f"   示例：第一条数据净利润 {sample/100000000:.2f}亿元")

            # 按报告期排序
            df = df.sort_values('end_date', ascending=False)
            print(f"   获取到数据条数: {len(df)}条")

            # 提取年份
            df['year'] = df['end_date'].str[:4].astype(int)
            print(f"   数据年份范围: {df['year'].min()}年 至 {df['year'].max()}年")

            # 按年份分组，每组优先使用年度数据
            yearly_data = {}

            # 确定年份范围：从最新年份往前推years年
            max_year = df['year'].max()
            print(f"   开始按年份处理数据（从{max_year}年往前推{years}年）...")

            for year in range(max_year, max_year - years - 1, -1):
                year_df = df[df['year'] == year]

                # 优先查找年度数据（1231结尾）
                annual_data = year_df[year_df['end_date'].str.endswith('1231')]

                if not annual_data.empty:
                    # 使用年度数据
                    row = annual_data.iloc[0]
                    # n_income_attr_p 已经是元，不需要转换
                    net_income_yuan = float(row['n_income_attr_p'])
                    yearly_data[year] = {
                        'end_date': row['end_date'],
                        'n_income_attr_p': net_income_yuan,
                        'source': f'{year}年(年度数据)'
                    }
                    print(f"   {year}年: 找到年度数据 {row['end_date']}，净利润 {net_income_yuan/100000000:.2f}亿元")
                else:
                    # 没有年度数据，使用该年最新的季度数据并年化
                    latest_quarter = year_df.iloc[0]  # 已经按日期降序排序
                    end_date_str = latest_quarter['end_date']
                    month_day = end_date_str[4:]  # 获取MMDD部分
                    quarterly_income = float(latest_quarter['n_income_attr_p'])  # 已经转换为元

                    print(f"   {year}年: 无年度数据，使用最新季度{end_date_str}，净利润{quarterly_income/100000000:.2f}亿元")

                    if pd.notna(quarterly_income) and quarterly_income > 0:
                        # 判断季度并年化
                        if month_day.endswith('0331'):  # Q1 一季度 (3个月)
                            annualized_income = quarterly_income * 4
                            source = f'{year}年(Q1年化{quarterly_income/100000000:.2f}亿×4)'
                        elif month_day.endswith('0630'):  # Q2 二季度 (6个月)
                            annualized_income = quarterly_income * 2
                            source = f'{year}年(Q2年化{quarterly_income/100000000:.2f}亿×2)'
                        elif month_day.endswith('0930'):  # Q3 三季度 (9个月)
                            annualized_income = quarterly_income * (4/3)
                            source = f'{year}年(Q3年化{quarterly_income/100000000:.2f}亿×4/3)'
                        elif month_day.endswith('1231'):  # Q4 四季度 (12个月，年度数据)
                            annualized_income = quarterly_income
                            source = f'{year}年(年度数据)'
                        else:
                            # 其他日期，保守估计不年化
                            annualized_income = quarterly_income
                            source = f'{year}年({month_day}数据)'

                        yearly_data[year] = {
                            'end_date': end_date_str,
                            'n_income_attr_p': annualized_income,
                            'source': source
                        }
                        print(f"         → 年化为: {annualized_income/100000000:.2f}亿元 ({source})")

            # 检查是否有足够数据
            if len(yearly_data) < 2:
                print(f"⚠️ 有效历史数据不足（需要至少2年，当前{len(yearly_data)}年）")
                return None

            # 按年份降序提取数据
            historical_incomes = []
            sources = []
            for year in sorted(yearly_data.keys(), reverse=True)[:years]:
                data = yearly_data[year]
                historical_incomes.append(data['n_income_attr_p'])
                sources.append(data['source'])

            print(f"✅ 获取历史净利润数据成功:")
            print(f"   注：数据已为n_income_attr_p（归属母公司净利润），单位为元")
            for income, source in zip(historical_incomes, sources):
                print(f"   {source}: {income/100000000:.2f}亿元")

            return historical_incomes

        except Exception as e:
            print(f"❌ 获取历史净利润数据失败: {e}")
            import traceback
            traceback.print_exc()
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

    def get_historical_fcf_for_dcf(self, max_years: int = 15) -> dict:
        """
        获取从上市以来的历史财务数据，用于DCF估值

        参数:
            max_years: 最多获取多少年的数据

        返回:
            包含历史FCF数据的字典
        """
        print(f"\n📊 获取 {self.ts_code} 的历史财务数据（用于DCF估值）...")

        try:
            # 获取上市日期（简化处理：从最早的交易日推算）
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=max_years*365)).strftime('%Y%m%d')

            print(f"   查询范围: {start_date} ~ {end_date}")

            # 获取利润表数据
            income_df = self.pro.income(
                ts_code=self.ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,end_date,end_type,update_flag,ann_date,revenue,operate_profit,total_profit,n_income_attr_p,'
                       'income_tax_exp,fin_exp,int_exp,fin_exp_int_exp'
            )

            # 获取资产负债表数据
            balance_df = self.pro.balancesheet(
                ts_code=self.ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,end_date,end_type,update_flag,total_assets,total_liab,total_hldr_eqy_exc_min_int,'
                       'money_cap,fix_assets,total_nca'
            )

            # 获取现金流量表数据
            # n_cashflow_act: 经营活动现金流
            # c_pay_acq_const_fiolta: 购建固定资产、无形资产和其他长期资产支付的现金（资本支出）
            cashflow_df = self.pro.cashflow(
                ts_code=self.ts_code,
                start_date=start_date,
                end_date=end_date,
                fields='ts_code,end_date,end_type,update_flag,n_cashflow_act,c_pay_acq_const_fiolta'
            )

            if income_df.empty:
                print(f"⚠️ 未获取到利润表数据")
                return None

            if balance_df.empty:
                print(f"⚠️ 未获取到资产负债表数据")
                return None

            # 筛选年报数据（只使用年报，跳过季报）
            # 使用end_type字段：4=年报, 3=三季报, 2=半年报, 1=一季报
            # 如果end_type字段不存在，回退到使用end_date以'1231'结尾筛选
            print(f"📊 财务报告数据筛选（只使用年报）:")

            # 利润表筛选
            if 'end_type' in income_df.columns:
                income_annual_only = income_df[income_df['end_type'] == '4'].copy()
                filter_method = "end_type=4"
            else:
                income_annual_only = income_df[income_df['end_date'].str.endswith('1231')].copy()
                filter_method = "end_date=1231"

            # 资产负债表筛选
            if 'end_type' in balance_df.columns:
                balance_annual_only = balance_df[balance_df['end_type'] == '4'].copy()
            else:
                balance_annual_only = balance_df[balance_df['end_date'].str.endswith('1231')].copy()

            # 去重：对于每个end_date，优先选择update_flag=1（最新版本），否则使用update_flag=0
            if 'update_flag' in income_annual_only.columns:
                # 按end_date分组，每组选择update_flag最大的记录
                income_annual_only = income_annual_only.sort_values('update_flag', ascending=False) \
                    .drop_duplicates(subset=['end_date'], keep='first')

            if 'update_flag' in balance_annual_only.columns:
                balance_annual_only = balance_annual_only.sort_values('update_flag', ascending=False) \
                    .drop_duplicates(subset=['end_date'], keep='first')

            income_annual_only['year'] = income_annual_only['end_date'].str[:4].astype(int)
            balance_annual_only['year'] = balance_annual_only['end_date'].str[:4].astype(int)

            print(f"   利润表: {len(income_df)} 份 → {len(income_annual_only)} 份年报（{filter_method}，已去重）")
            print(f"   资产负债表: {len(balance_df)} 份 → {len(balance_annual_only)} 份年报（{filter_method}，已去重）")

            # 合并利润表和资产负债表（只按year合并，避免end_date不一致导致重复）
            financial_data = income_annual_only.merge(
                balance_annual_only,
                on=['ts_code', 'year'],
                how='outer',
                suffixes=('', '_balance')
            )

            # 如果有现金流量表，也只使用年报数据
            if not cashflow_df.empty:
                # 筛选年报：end_type='4'（年报），如果不存在则使用end_date以'1231'结尾
                if 'end_type' in cashflow_df.columns:
                    cashflow_annual_only = cashflow_df[cashflow_df['end_type'] == '4'].copy()
                    cashflow_filter_method = "end_type=4"
                else:
                    cashflow_annual_only = cashflow_df[cashflow_df['end_date'].str.endswith('1231')].copy()
                    cashflow_filter_method = "end_date=1231"

                # 去重：优先选择update_flag=1（最新版本）
                if 'update_flag' in cashflow_annual_only.columns:
                    cashflow_annual_only = cashflow_annual_only.sort_values('update_flag', ascending=False) \
                        .drop_duplicates(subset=['end_date'], keep='first')

                cashflow_annual_only['year'] = cashflow_annual_only['end_date'].str[:4].astype(int)

                print(f"   现金流量表: {len(cashflow_df)} 份 → {len(cashflow_annual_only)} 份年报（{cashflow_filter_method}，已去重）")

                # 合并现金流量表（只按year合并）
                financial_data = financial_data.merge(
                    cashflow_annual_only,
                    on=['ts_code', 'year'],
                    how='outer',
                    suffixes=('', '_cashflow')
                )

            # 去重（如果仍有重复）
            financial_data = financial_data.drop_duplicates(subset=['ts_code', 'year']).reset_index(drop=True)

            # 按年份排序
            financial_data = financial_data.sort_values('year').reset_index(drop=True)

            # 计算自由现金流
            financial_data = self._calculate_fcf(financial_data)

            # 转换为字典格式返回
            historical_fcf = {
                'years': int(len(financial_data)),
                'year_range': [int(financial_data['year'].min()), int(financial_data['year'].max())],
                'data': []
            }

            for _, row in financial_data.iterrows():
                year_data = {
                    'year': int(row['year']),
                    'revenue': float(row.get('revenue', 0)) / 100000000,  # 转为亿元
                    'operate_profit': float(row.get('operate_profit', 0)) / 100000000,
                    'net_income': float(row.get('n_income_attr_p', 0)) / 100000000,
                    'nopat': float(row.get('nopat', 0)) / 100000000,
                    'depreciation': float(row.get('depreciation', 0)) / 100000000,
                    'capex': float(row.get('capex', 0)) / 100000000,
                    'wc_change': float(row.get('wc_change', 0)) / 100000000,
                    'fcf': float(row.get('fcf', 0)) / 100000000,
                }
                historical_fcf['data'].append(year_data)

            print(f"✅ 获取到 {len(financial_data)} 年的财务数据")
            print(f"   年份范围: {financial_data['year'].min()} ~ {financial_data['year'].max()}")

            # 显示最近几年数据
            print(f"\n📋 最近5年FCF数据:")
            print(f"{'年份':<6} {'营收(亿)':<12} {'NOPAT(亿)':<12} {'FCF(亿)':<12}")
            print("-"*50)
            for year_data in historical_fcf['data'][-5:]:
                print(f"{year_data['year']:<6} {year_data['revenue']:>10.2f}     "
                      f"{year_data['nopat']:>10.2f}     {year_data['fcf']:>10.2f}")

            return historical_fcf

        except Exception as e:
            print(f"❌ 获取历史财务数据失败: {e}")
            import traceback
            traceback.print_exc()
            return None

    def _calculate_fcf(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        计算自由现金流（内部方法）

        FCFF = 经营活动现金流 + 利息支出×(1-T) - 资本支出
        在中国会计准则下，利息支出归入经营活动现金流出，因此需加回税后利息。
        或者：FCF = NOPAT + 折旧摊销 - 资本支出 - 营运资本增加
        """
        # 计算实际有效税率：所得税费用 / 利润总额
        df['effective_tax_rate'] = 0.25  # 默认25%
        if 'income_tax_exp' in df.columns and 'total_profit' in df.columns:
            tax_exp = df['income_tax_exp'].fillna(0)
            total_profit = df['total_profit'].fillna(0)
            # 只在利润总额为正时计算有效税率
            valid = (total_profit > 0) & (tax_exp >= 0)
            df.loc[valid, 'effective_tax_rate'] = (tax_exp[valid] / total_profit[valid]).clip(0, 0.5)
        tax_rate = df['effective_tax_rate']

        # 方法1：优先使用现金流量表数据（最准确）
        if 'n_cashflow_act' in df.columns:
            # 经营活动现金流
            df['ocf'] = df['n_cashflow_act'].fillna(0)

            # 资本支出：购建固定资产、无形资产和其他长期资产支付的现金
            if 'c_pay_acq_const_fiolta' in df.columns:
                df['capex'] = df['c_pay_acq_const_fiolta'].fillna(0)
            else:
                if 'n_cash_flows_inv_act' in df.columns:
                    df['capex'] = df['n_cash_flows_inv_act'].fillna(0)
                else:
                    df['capex'] = 0

            # 利息支出：从利润表获取 fin_exp_int_exp（利息费用-利息支出）
            # 在中国准则下利息支出归入经营活动，需加回税后利息
            if 'fin_exp_int_exp' in df.columns:
                interest_exp = df['fin_exp_int_exp'].fillna(0)
            elif 'int_exp' in df.columns:
                interest_exp = df['int_exp'].fillna(0)
            else:
                interest_exp = 0
            after_tax_interest = interest_exp * (1 - tax_rate)

            # FCFF = OCF + 利息×(1-T) - CapEx
            df['fcf'] = df['ocf'] + after_tax_interest - df['capex']
            df['interest_exp'] = interest_exp
            df['after_tax_interest'] = after_tax_interest

            # NOPAT = 营业利润 × (1 - 有效税率)
            if 'operate_profit' in df.columns:
                df['nopat'] = df['operate_profit'].fillna(0) * (1 - tax_rate)
            else:
                df['nopat'] = df['n_income_attr_p'].fillna(0)

            if 'fix_assets' in df.columns:
                df['depreciation'] = df['fix_assets'] * 0.08
            else:
                df['depreciation'] = 0

            # wc_change设为0（因为使用OCF方法时不需要计算营运资本变化）
            df['wc_change'] = 0

        else:
            # 方法2：如果没有现金流数据，使用损益表和资产负债表估算
            # NOPAT (税后营业利润) - 使用有效税率
            if 'operate_profit' in df.columns:
                df['nopat'] = df['operate_profit'].fillna(0) * (1 - tax_rate)
            else:
                df['nopat'] = df['n_income_attr_p'].fillna(0)

            # 折旧摊销（简化处理）
            if 'fix_assets' in df.columns:
                df['depreciation'] = df['fix_assets'] * 0.08
            else:
                df['depreciation'] = 0

            # 资本支出（用固定资产变化估算）
            if 'fix_assets' in df.columns:
                df['fix_assets_lag'] = df['fix_assets'].shift(1)
                df['capex'] = (df['fix_assets'] - df['fix_assets_lag']).fillna(0)
                # 资本支出不能为负（至少为0）
                df['capex'] = df['capex'].apply(lambda x: max(x, 0))
            else:
                df['capex'] = 0

            # 营运资本变化（使用正确定义：流动资产 - 流动负债）
            if 'total_current_assets' in df.columns and 'total_current_liab' in df.columns:
                df['working_capital'] = df['total_current_assets'] - df['total_current_liab']
                df['wc_lag'] = df['working_capital'].shift(1)
                df['wc_change'] = df['working_capital'] - df['wc_lag']
                df['wc_change'] = df['wc_change'].fillna(0)
            else:
                df['wc_change'] = 0

            # FCF = NOPAT + 折旧摊销 - 资本支出 - 营运资本增加
            df['fcf'] = df['nopat'] + df['depreciation'] - df['capex'] - df['wc_change']
            df['ocf'] = df['fcf']  # 用于后续处理

        return df


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


def get_stock_industry_classification(stock_code):
    """
    获取股票的申万行业分类（使用与第二章相对估值相同的方法）

    参数:
        stock_code: 股票代码，如 '300735.SZ'

    返回:
        包含行业信息和行业指数代码的字典
    """
    try:
        pro = _init_tushare_pro()
        # 获取申万三级行业分类（与第二章相对估值一致）
        df_industry = pro.index_member_all(ts_code=stock_code)

        if df_industry.empty:
            print(f"⚠️ 未获取到{stock_code}的行业分类信息")
            return None

        df_industry = df_industry.sort_values('in_date', ascending=False)
        latest_industry = df_industry.iloc[0]

        sw_l1_code = latest_industry['l1_code']  # 申万一级行业代码
        sw_l1_name = latest_industry['l1_name']  # 申万一级行业名称
        sw_l2_code = latest_industry['l2_code']  # 申万二级行业代码
        sw_l2_name = latest_industry['l2_name']  # 申万二级行业名称
        sw_l3_code = latest_industry['l3_code']  # 申万三级行业代码
        sw_l3_name = latest_industry['l3_name']  # 申万三级行业名称

        print(f"✅ 获取到申万行业分类:")
        print(f"   一级: {sw_l1_name} ({sw_l1_code})")
        if sw_l2_name:
            print(f"   二级: {sw_l2_name} ({sw_l2_code})")
        if sw_l3_name:
            print(f"   三级: {sw_l3_name} ({sw_l3_code})")

        # 将申万三级行业代码转换为行业指数代码
        # 申万三级行业指数代码格式: 850531.SI (消费电子零部件及组装)
        # 注意：index_member_all返回的l3_code可能已经包含.SI后缀

        # 申万三级行业指数代码（检查是否已有.SI后缀）
        if sw_l3_code:
            sw_l3_index_code = sw_l3_code if sw_l3_code.endswith('.SI') else f"{sw_l3_code}.SI"
        else:
            sw_l3_index_code = sw_l1_code if sw_l1_code.endswith('.SI') else f"{sw_l1_code}.SI"

        result = {
            'sw_l1_code': sw_l1_code,
            'sw_l1_name': sw_l1_name,
            'sw_l2_code': sw_l2_code,
            'sw_l2_name': sw_l2_name,
            'sw_l3_code': sw_l3_code,
            'sw_l3_name': sw_l3_name,
            'sw_l3_index_code': sw_l3_index_code,  # 申万三级行业指数代码(用于获取历史数据)
        }

        print(f"   申万三级行业指数代码: {sw_l3_index_code}")

        return result

    except Exception as e:
        print(f"⚠️ 获取行业分类失败: {e}")
        import traceback
        traceback.print_exc()
        # 返回默认值（电子行业）
        return {
            'sw_l1_code': '801010',
            'sw_l1_name': '电子',
            'sw_l2_code': '',
            'sw_l2_name': '',
            'sw_l3_code': '',
            'sw_l3_name': '',
            'sw_l3_index_code': '801010.SI',  # 默认申万电子指数
        }


def fetch_industry_index_data(index_code, days=500):
    """
    获取申万行业指数历史数据（使用sw_daily接口）

    参数:
        index_code: 行业指数代码，如 '801010.SI' (申万电子指数) 或 '850531.SI' (三级行业指数)
        days: 获取天数

    返回:
        DataFrame或None
    """
    try:
        pro = _init_tushare_pro()
        # 计算日期范围
        end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.now() - timedelta(days=days*2)).strftime('%Y%m%d')

        print(f"   查询时间范围: {start_date} ~ {end_date}")

        # 使用申万行业日线行情接口（sw_daily）
        df = pro.sw_daily(ts_code=index_code, start_date=start_date, end_date=end_date)

        if df is None or df.empty:
            print(f"   ⚠️ 未获取到数据")
            return None

        # 按日期排序
        df = df.sort_values('trade_date')
        df.reset_index(drop=True, inplace=True)

        print(f"   ✅ 获取到 {len(df)} 条数据")
        print(f"   数据范围: {df['trade_date'].iloc[0]} ~ {df['trade_date'].iloc[-1]}")

        # sw_daily接口返回的字段名与index_daily相同，无需转换

        return df

    except Exception as e:
        print(f"   ❌ 获取{index_code}数据失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def generate_industry_data(stock_code, days=500):
    """
    生成行业指数数据（用于对比分析）

    参数:
        stock_code: 股票代码
        days: 获取天数

    返回:
        包含行业指数数据的字典
    """
    print("\n" + "="*70)
    print("📊 生成行业指数数据（申万行业分类）")
    print("="*70)

    # 获取行业分类
    industry_info = get_stock_industry_classification(stock_code)

    if not industry_info:
        print("⚠️ 无法获取行业分类，跳过行业数据生成")
        return None

    # 使用申万三级行业指数代码
    index_code = industry_info.get('sw_l3_index_code')

    # 行业名称显示（使用三级行业名称，如果没有则使用一级或二级）
    sw_l3_name = industry_info.get('sw_l3_name', '')
    sw_l2_name = industry_info.get('sw_l2_name', '')
    sw_l1_name = industry_info.get('sw_l1_name', '')

    if sw_l3_name:
        industry_display_name = sw_l3_name
    elif sw_l2_name:
        industry_display_name = f"{sw_l1_name}-{sw_l2_name}"
    else:
        industry_display_name = sw_l1_name

    if not index_code:
        print("⚠️ 未找到行业指数代码")
        return None

    print(f"📡 正在获取 {industry_display_name}（{index_code}）的行业指数数据...")

    # 获取行业指数历史数据
    df = fetch_industry_index_data(index_code, days=days)

    if df is None or len(df) < 250:
        print(f"⚠️ 行业指数数据不足，跳过")
        return None

    # 计算行业指数指标（使用与个股相同的计算方法）
    latest_date = df['trade_date'].iloc[-1]
    current_level = df['close'].iloc[-1]

    # 波动率
    vol_20 = calculate_rolling_volatility(df, 20)
    vol_60 = calculate_rolling_volatility(df, 60)
    vol_120 = calculate_rolling_volatility(df, 120)
    vol_250 = calculate_rolling_volatility(df, 250)

    # 年化收益率
    annual_return_20d = calculate_annual_return(df, 20)
    annual_return_60d = calculate_annual_return(df, 60)
    annual_return_120d = calculate_annual_return(df, 120)
    annual_return_250d = calculate_annual_return(df, 250)

    # 区间收益率
    period_return_20d = calculate_period_return(df, 20)
    period_return_60d = calculate_period_return(df, 60)
    period_return_120d = calculate_period_return(df, 120)
    period_return_250d = calculate_period_return(df, 250)

    # 移动平均线
    ma_20 = df['close'].rolling(window=20).mean().iloc[-1]
    ma_60 = df['close'].rolling(window=60).mean().iloc[-1]
    ma_120 = df['close'].rolling(window=120).mean().iloc[-1]
    ma_250 = df['close'].rolling(window=250).mean().iloc[-1]

    # 胜率（兼容pct_chg和pct_change字段，多个时间窗口）
    pct_col = 'pct_chg' if 'pct_chg' in df.columns else 'pct_change'
    win_rate_20d = (df[pct_col].iloc[-20:] > 0).sum() / 20 if len(df) >= 20 else 0
    win_rate_60d = (df[pct_col].iloc[-60:] > 0).sum() / 60 if len(df) >= 60 else 0
    win_rate_120d = (df[pct_col].iloc[-120:] > 0).sum() / 120 if len(df) >= 120 else 0
    win_rate_250d = (df[pct_col].iloc[-250:] > 0).sum() / 250 if len(df) >= 250 else 0

    # 构建行业数据字典
    industry_data = {
        'index_code': index_code,
        'industry_name': industry_display_name,
        'sw_l1_code': industry_info.get('sw_l1_code', ''),
        'sw_l1_name': sw_l1_name,
        'sw_l2_code': industry_info.get('sw_l2_code', ''),
        'sw_l2_name': sw_l2_name,
        'sw_l3_code': industry_info.get('sw_l3_code', ''),
        'sw_l3_name': sw_l3_name,
        'analysis_date': latest_date,
        'current_level': round(float(current_level), 2),
        'volatility_20d': round(float(vol_20['latest']), 4),
        'volatility_60d': round(float(vol_60['latest']), 4),
        'volatility_120d': round(float(vol_120['latest']), 4),
        'volatility_250d': round(float(vol_250['latest']), 4),
        'volatility': round(float(vol_60['latest']), 4),  # 默认60日
        'annual_return_20d': round(float(annual_return_20d), 4),
        'annual_return_60d': round(float(annual_return_60d), 4),
        'annual_return_120d': round(float(annual_return_120d), 4),
        'annual_return_250d': round(float(annual_return_250d), 4),
        'period_return_20d': round(float(period_return_20d), 4),
        'period_return_60d': round(float(period_return_60d), 4),
        'period_return_120d': round(float(period_return_120d), 4),
        'period_return_250d': round(float(period_return_250d), 4),
        'drift': round(float(annual_return_60d), 4),
        'ma_20': round(float(ma_20), 2),
        'ma_60': round(float(ma_60), 2),
        'ma_120': round(float(ma_120), 2),
        'ma_250': round(float(ma_250), 2),
        'win_rate_20d': round(float(win_rate_20d), 4),
        'win_rate_60d': round(float(win_rate_60d), 4),
        'win_rate_120d': round(float(win_rate_120d), 4),
        'win_rate_250d': round(float(win_rate_250d), 4),
        'total_days': len(df),
        'data_source': 'tushare_sw_index',
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    return industry_data


def update_placement_params_with_fcf(stock_code, ma20, data_dir):
    """
    更新placement_params.json，添加历史FCF数据

    参数:
        stock_code: 股票代码
        ma20: MA20价格
        data_dir: 数据目录
    """
    try:
        placement_file = os.path.join(data_dir, f"{stock_code.replace('.', '_')}_placement_params.json")

        if not os.path.exists(placement_file):
            print(f"⚠️ 未找到 {placement_file}，跳过FCF数据更新")
            return

        # 读取placement_params.json
        with open(placement_file, 'r', encoding='utf-8') as f:
            placement_params = json.load(f)

        # 添加历史FCF数据（使用修复后的FCF计算方法）
        print(f"📊 获取历史FCF数据...")
        financial = TushareFinancialData(stock_code)
        historical_fcf = financial.get_historical_fcf_for_dcf(max_years=15)

        if historical_fcf and historical_fcf.get('data'):
            placement_params['historical_fcf_data'] = historical_fcf
            print(f"✅ 已添加 {len(historical_fcf['data'])} 年历史FCF数据")

            # 显示最近几年的FCF
            print(f"   最近几年FCF（亿元）:")
            for item in historical_fcf['data'][-5:]:
                print(f"   {item['year']}: {item['fcf']:.2f}")
        else:
            print(f"⚠️ 未获取到历史FCF数据")

        # 保存
        with open(placement_file, 'w', encoding='utf-8') as f:
            json.dump(placement_params, f, indent=2, ensure_ascii=False)

        print(f"✅ 已更新FCF数据到: {placement_file}")

    except Exception as e:
        print(f"⚠️ 更新FCF数据失败: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='生成最新市场数据')
    parser.add_argument('--stock', type=str, default='300735.SZ', help='股票代码（默认：300735.SZ）')
    parser.add_argument('--name', type=str, default=None, help='股票名称（可选）')
    args = parser.parse_args()

    stock_code = args.stock
    stock_name = args.name if args.name else stock_code  # 如果没有提供名称，使用股票代码

    print("="*70)
    print("🚀 生成最新市场数据（使用当前日期回溯）")
    print("="*70)

    # 生成数据
    market_data = generate_market_data(stock_code, stock_name)

    if market_data:
        # 打印摘要
        print_market_data_summary(market_data)

        # 保存文件到项目根目录的data目录
        filename = f"{stock_code.replace('.', '_')}_market_data.json"

        # 获取脚本所在目录的绝对路径
        script_dir = os.path.dirname(os.path.abspath(__file__))
        # data目录是项目根目录下的data文件夹
        data_dir = os.path.join(os.path.dirname(script_dir), 'data')

        # 确保data目录存在
        os.makedirs(data_dir, exist_ok=True)

        # 保存到data目录
        filepath = os.path.join(data_dir, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(market_data, f, ensure_ascii=False, indent=2)
        print(f"\n✅ 已保存市场数据到: {filepath}")

        # 注意：不再更新placement_params.json中的issue_price
        # 因为新系统会根据MA20和premium_rate自动计算issue_price
        # update_placement_params_issue_price(stock_code, market_data['ma_20'], data_dir)

        # 添加历史FCF数据到placement_params.json
        print("\n" + "="*70)
        print("📊 添加历史FCF数据到 placement_params.json")
        print("="*70)

        placement_file = os.path.join(data_dir, f"{stock_code.replace('.', '_')}_placement_params.json")

        # 如果文件不存在，创建默认配置
        if not os.path.exists(placement_file):
            print(f"⚠️ {placement_file} 不存在，创建默认配置文件")

            # 创建默认配置（新格式，不包含issue_price和current_price）
            placement_params = {
                "stock_code": stock_code,
                "stock_name": stock_name,
                "financing_amount": 0.0,
                "lockup_period": 6,
                "pricing_method": "ma20_discount_90",
                "premium_rate": -0.10,
                "risk_free_rate": 0.03,
                "net_assets": 0.0,
                "total_debt": 0.0,
                "net_income": 0.0,
                "revenue_growth": 0.15,
                "operating_margin": 0.15,
                "beta": 1.0,
                "historical_fcf_data": {
                    "years": 5,
                    "year_range": [2020, 2024],
                    "data": []
                },
                "_notes": {
                    "financing_amount": "融资金额（元）- 必填，请在定增公告中查找",
                    "lockup_period": "锁定期（月）- 默认6个月",
                    "pricing_method": "定价方式：ma20_discount_90(MA20九折), ma20_discount_85(MA20八五折), ma20_par(MA20平价), custom_premium(自定义溢价率)",
                    "premium_rate": "溢价率（负数为折价，正数为溢价）- 默认-0.10表示九折",
                    "_auto_generated": "以下参数自动计算，无需手动填写",
                    "issue_price": f"自动计算：MA20({market_data.get('ma_20', 0):.2f}) × (1 + premium_rate)",
                    "current_price": f"自动从API获取最新股价: {market_data.get('current_price', 0):.2f}元"
                },
                "created_at": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "updated_at": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }

            # 保存默认配置
            with open(placement_file, 'w', encoding='utf-8') as f:
                json.dump(placement_params, f, indent=2, ensure_ascii=False)

            print(f"✅ 已创建默认配置文件: {placement_file}")
            print(f"   MA20: {market_data.get('ma_20', 0):.2f} 元")
            print(f"   当前价: {market_data.get('current_price', 0):.2f} 元")
        else:
            # 读取现有配置
            with open(placement_file, 'r', encoding='utf-8') as f:
                placement_params = json.load(f)

            # 获取历史FCF数据
            print(f"正在获取历史FCF数据...")
            financial = TushareFinancialData(stock_code)
            historical_fcf = financial.get_historical_fcf_for_dcf(max_years=15)

            if historical_fcf and historical_fcf.get('data'):
                placement_params['historical_fcf_data'] = historical_fcf
                print(f"✅ 已添加 {len(historical_fcf['data'])} 年历史FCF数据")

                # 显示最近几年的FCF
                print(f"   最近几年FCF（亿元）:")
                for item in historical_fcf['data'][-5:]:
                    print(f"   {item['year']}: {item['fcf']:.2f}")

                # 保存更新后的文件
                with open(placement_file, 'w', encoding='utf-8') as f:
                    json.dump(placement_params, f, indent=2, ensure_ascii=False)
                print(f"✅ 已更新FCF数据到: {placement_file}")
            else:
                print(f"⚠️ 未获取到历史FCF数据")

        # 生成行业指数数据
        print("\n" + "="*70)
        print("📊 生成行业指数数据")
        print("="*70)

        industry_data = generate_industry_data(stock_code)

        if industry_data:
            # 保存行业数据
            industry_filename = f"{stock_code.replace('.', '_')}_industry_data.json"
            industry_filepath = os.path.join(data_dir, industry_filename)

            with open(industry_filepath, 'w', encoding='utf-8') as f:
                json.dump(industry_data, f, ensure_ascii=False, indent=2)

            print(f"\n✅ 已保存行业指数数据到: {industry_filepath}")

            # 打印行业数据摘要
            print("\n📊 行业指数数据摘要:")
            print(f"   行业: {industry_data['sw_l1_name']}", end='')
            if industry_data.get('sw_l2_name'):
                print(f" -> {industry_data['sw_l2_name']}", end='')
            if industry_data.get('sw_l3_name'):
                print(f" -> {industry_data['sw_l3_name']}")
            else:
                print()
            print(f"   指数代码: {industry_data['index_code']}")
            print(f"   当前点位: {industry_data['current_level']:.2f}")
            print(f"   波动率(60日): {industry_data['volatility_60d']*100:.2f}%")
            print(f"   年化收益率(60日): {industry_data['annual_return_60d']*100:+.2f}% (连续复利)")
            print(f"   区间收益率(60日): {industry_data['period_return_60d']*100:+.2f}% (连续复利)")
            print(f"   年化收益率(250日): {industry_data['annual_return_250d']*100:+.2f}% (连续复利)")
            print(f"   区间收益率(250日): {industry_data['period_return_250d']*100:+.2f}% (连续复利)")

        print("\n📌 使用方法:")
        print(f"   from utils.market_data_loader import load_market_data")
        print(f"   market_data = load_market_data('{stock_code}')")
        print(f"   industry_data = load_market_data('{stock_code}', data_type='industry')")
