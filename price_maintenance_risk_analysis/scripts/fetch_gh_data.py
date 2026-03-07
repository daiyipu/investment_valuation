#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
股票定增分析数据获取脚本（通用版本）
从 Tushare 获取实时数据并配置定增分析参数

支持的股票：
- 光弘科技 300735.SZ
- 其他A股：修改 STOCK_CODE 配置即可
"""

import sys
import os
sys.path.append('..')

import pandas as pd
import tushare as ts
from datetime import datetime, timedelta


def get_stock_data(ts_code, token=None):
    """
    获取指定股票的实时数据（通用函数）

    参数:
        ts_code: 股票代码，如 '300735.SZ' 或 '000001.SZ'
        token: Tushare API Token，如果为 None 则从环境变量读取

    返回:
        包含股票数据的字典
    """
    # 获取 token
    if token is None:
        token = os.environ.get('TUSHARE_TOKEN')
        if not token:
            print("⚠️ 请设置 TUSHARE_TOKEN 环境变量或传入 token 参数")
            print("   方法1: export TUSHARE_TOKEN='your_token'")
            print("   方法2: get_stock_data('300735.SZ', token='your_token')")
            return None

    # 初始化 Tushare API
    pro = ts.pro_api(token)

    print(f"正在获取 {ts_code} 的数据...")

    result = {}

    try:
        # 1. 获取基本信息
        df_basic = pro.stock_basic(
            ts_code=ts_code,
            fields='ts_code,symbol,name,area,industry,list_date,market'
        )

        if df_basic.empty:
            print(f"❌ 未找到股票 {ts_code}")
            return None

        basic_info = df_basic.iloc[0]
        result['name'] = basic_info['name']
        result['ts_code'] = ts_code
        result['industry'] = basic_info['industry']
        result['list_date'] = basic_info['list_date']
        result['market'] = basic_info['market']

        print(f"✅ 公司: {result['name']}")
        print(f"   行业: {result['industry']}")

        # 2. 获取最新交易日
        end_date = datetime.now()
        start_date = end_date - timedelta(days=10)
        df_cal = pro.trade_cal(
            exchange='SZSE',
            start_date=start_date.strftime('%Y%m%d'),
            end_date=end_date.strftime('%Y%m%d'),
            fields='cal_date,is_open'
        )

        if not df_cal.empty:
            trade_days = df_cal[df_cal['is_open'] == 1]['cal_date'].tolist()
            if trade_days:
                trade_date = trade_days[-1]
            else:
                trade_date = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
        else:
            trade_date = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')

        print(f"   交易日: {trade_date}")

        # 3. 获取最新行情
        df_daily = pro.daily(
            ts_code=ts_code,
            trade_date=trade_date,
            fields='ts_code,trade_date,open,high,low,close,pre_close,vol,amount,pct_chg'
        )

        if df_daily.empty:
            # 尝试获取最近几天的数据
            df_daily = pro.daily(
                ts_code=ts_code,
                start_date=(datetime.now() - timedelta(days=5)).strftime('%Y%m%d'),
                end_date=datetime.now().strftime('%Y%m%d'),
                fields='ts_code,trade_date,open,high,low,close,pre_close,vol,amount,pct_chg'
            )

        if not df_daily.empty:
            latest = df_daily.iloc[-1]
            result['current_price'] = float(latest['close'])
            result['trade_date'] = latest['trade_date']
            result['daily_change'] = float(latest['pct_chg'])
            print(f"   当前价: {result['current_price']:.2f} 元")
            print(f"   日涨跌: {result['daily_change']:.2f}%")

        # 4. 获取日线基础数据（估值指标）
        df_basic_data = pro.daily_basic(
            ts_code=ts_code,
            trade_date=trade_date,
            fields='ts_code,trade_date,pe_ttm,pe,ps_ttm,ps,pb,total_mv,circ_mv'
        )

        if df_basic_data.empty:
            df_basic_data = pro.daily_basic(
                ts_code=ts_code,
                start_date=(datetime.now() - timedelta(days=5)).strftime('%Y%m%d'),
                end_date=datetime.now().strftime('%Y%m%d'),
                fields='ts_code,trade_date,pe_ttm,pe,ps_ttm,ps,pb,total_mv,circ_mv'
            )

        if not df_basic_data.empty:
            latest_basic = df_basic_data.iloc[-1]
            result['pe_ttm'] = float(latest_basic['pe_ttm']) if pd.notna(latest_basic['pe_ttm']) else None
            result['pb'] = float(latest_basic['pb']) if pd.notna(latest_basic['pb']) else None
            result['market_cap'] = float(latest_basic['total_mv']) if pd.notna(latest_basic['total_mv']) else None
            print(f"   市盈率(TTM): {result['pe_ttm']:.2f}" if result['pe_ttm'] else "   市盈率: N/A")
            print(f"   市净率: {result['pb']:.2f}" if result['pb'] else "   市净率: N/A")
            print(f"   总市值: {result['market_cap']:.2f} 亿元" if result['market_cap'] else "   总市值: N/A")

        # 5. 获取财务数据
        df_income = pro.income(
            ts_code=ts_code,
            limit=1,
            fields='ts_code,ann_date,end_date,revenue,operate_profit,total_profit,n_income'
        )

        if not df_income.empty:
            income = df_income.iloc[0]
            result['revenue'] = float(income['revenue']) if pd.notna(income['revenue']) else 0
            result['net_income'] = float(income['n_income']) if pd.notna(income['n_income']) else 0
            result['report_date'] = income['end_date']
            print(f"   营业收入: {result['revenue']/100000000:.2f} 亿元" if result['revenue'] else "   营业收入: N/A")
            print(f"   净利润: {result['net_income']/100000000:.2f} 亿元" if result['net_income'] else "   净利润: N/A")

        # 6. 获取资产负债数据
        df_balance = pro.balancesheet(
            ts_code=ts_code,
            limit=1,
            fields='ts_code,end_date,total_assets,total_hldr_eqy_exc_min_int,total_liab'
        )

        if not df_balance.empty:
            balance = df_balance.iloc[0]
            result['total_assets'] = float(balance['total_assets']) if pd.notna(balance['total_assets']) else 0
            result['net_assets'] = float(balance['total_hldr_eqy_exc_min_int']) if pd.notna(balance['total_hldr_eqy_exc_min_int']) else 0
            result['total_debt'] = float(balance['total_liab']) if pd.notna(balance['total_liab']) else 0
            print(f"   总资产: {result['total_assets']/100000000:.2f} 亿元")
            print(f"   净资产: {result['net_assets']/100000000:.2f} 亿元")

        # 7. 获取历史波动率（基于过去60个交易日）
        start_date_60 = (datetime.now() - timedelta(days=90)).strftime('%Y%m%d')
        df_history = pro.daily(
            ts_code=ts_code,
            start_date=start_date_60,
            end_date=datetime.now().strftime('%Y%m%d'),
            fields='ts_code,trade_date,close,pct_chg'
        )

        if not df_history.empty and len(df_history) > 20:
            df_history = df_history.sort_values('trade_date')
            # 计算历史波动率（年化）
            volatility = df_history['pct_chg'].std() * (252 ** 0.5) / 100
            result['volatility'] = volatility
            print(f"   历史波动率(60日): {volatility*100:.2f}%")

        return result

    except Exception as e:
        print(f"❌ 获取数据失败: {e}")
        import traceback
        traceback.print_exc()
        return None


def get_real_placement_price(ts_code, pro):
    """
    获取真实的定增发行价格

    参数:
        ts_code: 股票代码
        pro: Tushare API对象

    返回:
        真实发行价，如果没有则返回None
    """
    try:
        # 尝试使用Tushare的定增接口
        df_placement = pro.issuance(
            ts_code=ts_code,
            fields='ts_code,ann_date,price_amount'
        )

        if not df_placement.empty:
            # 获取最新的定增数据
            latest_placement = df_placement.iloc[-1]
            real_price = float(latest_placement['price_amount'])
            return real_price

        return None
    except Exception as e:
        # 如果接口调用失败或没有数据，返回None
        return None


def calculate_ma30_price(ts_code, pro):
    """
    计算MA30价格（30日移动平均线）

    参数:
        ts_code: 股票代码
        pro: Tushare API对象

    返回:
        MA30价格
    """
    try:
        # 获取过去90天的数据（确保有30个交易日）
        end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.now() - timedelta(days=90)).strftime('%Y%m%d')

        df_history = pro.daily(
            ts_code=ts_code,
            start_date=start_date,
            end_date=end_date,
            fields='ts_code,trade_date,close'
        )

        if df_history.empty or len(df_history) < 30:
            return None

        df_history = df_history.sort_values('trade_date')
        ma30 = df_history['close'].rolling(window=30).mean().iloc[-1]

        return float(ma30)
    except Exception as e:
        return None


def generate_private_placement_params(stock_data, custom_params=None):
    """
    根据股票数据生成定增分析参数

    参数:
        stock_data: 从 get_stock_data 获取的股票数据
        custom_params: 自定义参数字典，覆盖默认值

    返回:
        定增分析参数字典
    """
    if stock_data is None:
        return None

    # 初始化Tushare API（用于获取真实定增价或MA30）
    try:
        pro = ts.pro_api(os.environ.get('TUSHARE_TOKEN'))
    except:
        pro = None

    # 确定发行价格
    issue_price = None
    price_source = "估算"

    if pro:
        # 1. 尝试获取真实定增发行价
        print(f"\n🔍 查找定增发行价...")
        real_price = get_real_placement_price(stock_data['ts_code'], pro)
        if real_price:
            issue_price = real_price
            price_source = "真实定增价"
            print(f"   ✅ 使用真实定增发行价: {issue_price:.2f} 元")
        else:
            # 2. 使用MA30的8折
            print(f"   ⚠️ 未找到真实定增价，尝试使用MA30的8折...")
            ma30 = calculate_ma30_price(stock_data['ts_code'], pro)
            if ma30:
                issue_price = ma30 * 0.8
                price_source = "MA30的8折"
                print(f"   MA30: {ma30:.2f} 元")
                print(f"   发行价 = MA30 × 0.8 = {issue_price:.2f} 元")
            else:
                # 3. 回退到当前价的8折
                issue_price = stock_data.get('current_price', 20) * 0.8
                price_source = "当前价的8折"
                print(f"   ⚠️ 无法计算MA30，使用当前价的8折: {issue_price:.2f} 元")
    else:
        # 无法连接API，使用当前价的8折
        issue_price = stock_data.get('current_price', 20) * 0.8
        price_source = "当前价的8折"
        print(f"   ⚠️ 使用当前价的8折: {issue_price:.2f} 元")

    # 默认定增参数（可自定义）
    default_params = {
        # 定增基本信息
        'issue_price': issue_price,  # 发行价（真实或估算）
        'price_source': price_source,  # 发行价来源
        'issue_shares': 5000000,  # 发行数量（股）
        'lockup_period': 12,  # 锁定期（月）
        'current_price': stock_data.get('current_price', 20),
        'risk_free_rate': 0.03,  # 无风险利率

        # 公司财务数据
        'net_assets': stock_data.get('net_assets', 5000000000),
        'total_debt': stock_data.get('total_debt', 2000000000),
        'net_income': stock_data.get('net_income', 800000000),

        # 增长假设
        'revenue_growth': 0.25,  # 营收增长率
        'operating_margin': 0.20,  # 营业利润率

        # 风险参数
        'volatility': stock_data.get('volatility', 0.35),  # 波动率
        'beta': 1.2,  # Beta系数
    }

    # 合并自定义参数
    if custom_params:
        default_params.update(custom_params)

    return default_params


def print_analysis_summary(stock_data, placement_params):
    """打印分析摘要"""
    print("\n" + "="*70)
    print(f"{stock_data.get('name', 'N/A')}定增分析参数")
    print("="*70)

    print(f"\n📊 公司基本信息:")
    print(f"   公司名称: {stock_data.get('name', 'N/A')}")
    print(f"   股票代码: {stock_data.get('ts_code', 'N/A')}")
    print(f"   所属行业: {stock_data.get('industry', 'N/A')}")
    print(f"   当前价格: {stock_data.get('current_price', 'N/A')} 元/股")
    print(f"   数据日期: {stock_data.get('trade_date', 'N/A')}")

    print(f"\n💰 定增发行参数:")
    print(f"   发行价格: {placement_params['issue_price']:.2f} 元/股")
    if 'price_source' in placement_params:
        print(f"   价格来源: {placement_params['price_source']}")
    print(f"   发行数量: {placement_params['issue_shares']:,} 股")
    print(f"   融资金额: {placement_params['issue_price'] * placement_params['issue_shares'] / 100000000:.2f} 亿元")
    print(f"   锁定期: {placement_params['lockup_period']} 个月")
    print(f"   发行折价: {(placement_params['current_price']/placement_params['issue_price'] - 1)*100:.2f}%")

    print(f"\n📈 财务与风险参数:")
    print(f"   净资产: {placement_params['net_assets']/100000000:.2f} 亿元")
    print(f"   总债务: {placement_params['total_debt']/100000000:.2f} 亿元")
    print(f"   净利润: {placement_params['net_income']/100000000:.2f} 亿元")
    print(f"   波动率: {placement_params['volatility']*100:.2f}%")

    print("="*70)


# ===== 股票配置 =====
# 修改这里来分析不同的股票
STOCK_CODE = '300735.SZ'  # 光弘科技
# STOCK_CODE = '000001.SZ'  # 平安银行
# STOCK_CODE = '600519.SH'  # 贵州茅台


if __name__ == "__main__":
    # 获取数据（需要设置 TUSHARE_TOKEN 环境变量）
    print("股票定增分析数据获取")
    print("="*70)
    print(f"目标股票: {STOCK_CODE}")

    # 检查环境变量
    token = os.environ.get('TUSHARE_TOKEN')
    if not token:
        print("\n⚠️ 请先设置 Tushare Token:")
        print("   export TUSHARE_TOKEN='your_token_here'")
        print("\n或者修改脚本传入 token:")
        print(f"   stock_data = get_stock_data('{STOCK_CODE}', token='your_token_here')")
        sys.exit(1)

    # 获取股票数据
    stock_data = get_stock_data(STOCK_CODE)

    if stock_data:
        # 生成定增参数
        placement_params = generate_private_placement_params(stock_data)

        # 打印摘要
        print_analysis_summary(stock_data, placement_params)

        # 保存参数到文件（供 notebook 使用）
        import json
        params_file = f'{STOCK_CODE.replace(".", "_")}_placement_params.json'
        with open(params_file, 'w', encoding='utf-8') as f:
            # 转换 numpy 类型为 Python 原生类型
            params_to_save = {}
            for k, v in placement_params.items():
                if isinstance(v, (int, float, str, bool, type(None))):
                    params_to_save[k] = v
                else:
                    params_to_save[k] = float(v) if hasattr(v, '__float__') else str(v)
            json.dump(params_to_save, f, indent=2, ensure_ascii=False)

        print(f"\n✅ 参数已保存到 {params_file}")
        print("   在 notebook 中使用以下代码加载:")
        print(f"   with open('{params_file}', 'r') as f:")
        print("       params = json.load(f)")
