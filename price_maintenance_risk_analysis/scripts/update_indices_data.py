#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
市场指数数据更新脚本（合并版）

功能：
1. 从Tushare API重新获取并计算指数数据（使用修复公式）
2. 补充缺失的窗口期指标

用法：
    python update_indices_data.py --mode rebuild    # 重新获取数据
    python update_indices_data.py --mode update     # 更新现有文件
    python update_indices_data.py                 # 默认rebuild模式
"""

import json
import os
import numpy as np
from datetime import datetime, timedelta
import argparse


def calculate_annual_return_v2(df, window):
    """
    计算年化收益率和区间对数收益率（连续复利）

    年化方法：年化收益率 = 区间对数收益率 × (250 / 窗口期天数)
    与第5、6章蒙特卡洛和情景分析保持一致

    返回：
        annual_log_return: 年化对数收益率
        period_log_return: 区间对数收益率（直接用于第6章漂移率）
    """
    if len(df) < window:
        return np.nan, np.nan

    # 获取window天前的收盘价和最新收盘价
    start_price = df['close'].iloc[-window]
    end_price = df['close'].iloc[-1]

    # 计算期间对数收益率（连续复利）
    period_log_return = np.log(end_price / start_price)

    # 按交易日比例年化（250个交易日/年）
    annual_log_return = period_log_return * (250.0 / window)

    return annual_log_return, period_log_return


def calculate_volatility(df, window):
    """计算波动率（年化）"""
    if len(df) < window:
        return np.nan

    # 计算对数收益率
    close_prices = df['close'].iloc[-window:].values
    log_returns = np.diff(np.log(close_prices))

    # 年化波动率（250个交易日）
    volatility = log_returns.std() * np.sqrt(250)

    return volatility


def fetch_index_data(ts_code, days=500, end_date=None):
    """
    获取指数历史数据

    参数:
        ts_code: 指数代码（如 '000001.SH'）
        days: 获取天数
        end_date: 结束日期（YYYYMMDD格式），如果为None则使用最新日期

    返回:
        DataFrame或None
    """
    try:
        import tushare as ts

        ts_token = os.environ.get('TUSHARE_TOKEN', '')
        if not ts_token:
            print("⚠️ 未设置TUSHARE_TOKEN环境变量")
            print("   请设置: export TUSHARE_TOKEN='your_token'")
            return None

        pro = ts.pro_api(ts_token)

        # 计算日期范围
        if end_date is None:
            end_date = datetime.now().strftime('%Y%m%d')
        start_date = (datetime.strptime(end_date, '%Y%m%d') - timedelta(days=days)).strftime('%Y%m%d')

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


def rebuild_locked_indices_data(stock_code, issue_date):
    """
    为特定项目生成基于发行日的锁定指数数据

    参数:
        stock_code: 股票代码（如 '300735.SZ'）
        issue_date: 发行日期（YYYYMMDD格式）

    返回:
        指标数据字典或None
    """
    # 计算发行日前一日作为数据截止日
    issue_date_obj = datetime.strptime(issue_date, '%Y%m%d')
    data_end_date = (issue_date_obj - timedelta(days=1)).strftime('%Y%m%d')

    print(f"🔄 为项目 {stock_code} 生成锁定指数数据（截至{data_end_date}，发行日{issue_date}）")

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

    results = {}

    for index_info in indices:
        name = index_info['name']
        code = index_info['code']

        print(f"📊 获取{name}数据（截至{data_end_date}）...")

        # 获取历史数据（指定截止日期）
        df = fetch_index_data(code, days=500, end_date=data_end_date)

        if df is None or len(df) < 250:
            print(f"  ⚠️ 数据不足，跳过")
            continue

        # 计算各窗口指标
        try:
            # 当前收盘价（截止日的收盘价）
            current_close = df['close'].iloc[-1]

            # 波动率（所有窗口期）
            vol_20d = calculate_volatility(df, 20)
            vol_60d = calculate_volatility(df, 60)
            vol_120d = calculate_volatility(df, 120)
            vol_250d = calculate_volatility(df, 250)

            # 年化收益率和区间对数收益率（所有窗口期）
            ann_20d, period_20d = calculate_annual_return_v2(df, 20)
            ann_60d, period_60d = calculate_annual_return_v2(df, 60)
            ann_120d, period_120d = calculate_annual_return_v2(df, 120)
            ann_250d, period_250d = calculate_annual_return_v2(df, 250)

            # 计算MA
            ma_20 = df['close'].iloc[-20:].mean()
            ma_60 = df['close'].iloc[-60:].mean()
            ma_120 = df['close'].iloc[-120:].mean()
            ma_250 = df['close'].iloc[-250:].mean()

            # 计算胜率
            recent_20 = df['close'].iloc[-20:]
            win_days_20 = (recent_20.diff() > 0).sum()
            win_rate_20d = win_days_20 / 20 if len(recent_20) >= 2 else 0.5

            recent_60 = df['close'].iloc[-60:]
            win_days_60 = (recent_60.diff() > 0).sum()
            win_rate_60d = win_days_60 / 60 if len(recent_60) >= 2 else 0.5

            recent_120 = df['close'].iloc[-120:]
            win_days_120 = (recent_120.diff() > 0).sum()
            win_rate_120d = win_days_120 / 120 if len(recent_120) >= 2 else 0.5

            recent_250 = df['close'].iloc[-250:]
            win_days_250 = (recent_250.diff() > 0).sum()
            win_rate_250d = win_days_250 / 250 if len(recent_250) >= 2 else 0.5

            # 构建结果（包含区间对数收益率）
            results[name] = {
                'current_level': current_close,
                # 波动率
                'volatility_20d': vol_20d if not np.isnan(vol_20d) else 0.30,
                'volatility_60d': vol_60d if not np.isnan(vol_60d) else 0.30,
                'volatility_120d': vol_120d if not np.isnan(vol_120d) else 0.35,
                'volatility_250d': vol_250d if not np.isnan(vol_250d) else 0.35,
                # 年化对数收益率
                'return_20d': ann_20d if not np.isnan(ann_20d) else 0,
                'return_60d': ann_60d if not np.isnan(ann_60d) else 0,
                'return_120d': ann_120d if not np.isnan(ann_120d) else 0,
                'return_250d': ann_250d if not np.isnan(ann_250d) else 0,
                # 区间对数收益率（新增字段）
                'period_log_return_20d': period_20d if not np.isnan(period_20d) else 0,
                'period_log_return_60d': period_60d if not np.isnan(period_60d) else 0,
                'period_log_return_120d': period_120d if not np.isnan(period_120d) else 0,
                'period_log_return_250d': period_250d if not np.isnan(period_250d) else 0,
                # 移动平均线
                'ma_20': ma_20,
                'ma_60': ma_60,
                'ma_120': ma_120,
                'ma_250': ma_250,
                # 胜率
                'win_rate_20d': win_rate_20d,
                'win_rate_60d': win_rate_60d,
                'win_rate_120d': win_rate_120d,
                'win_rate_250d': win_rate_250d,
                # 锁定信息
                'locked_issue_date': issue_date,
                'locked_data_date': data_end_date,
            }

            print(f"  ✅ 完成: 120日区间收益率={period_120d*100:+.2f}%")

        except Exception as e:
            print(f"  ❌ 计算失败: {e}")
            continue

    # 保存锁定数据（覆盖通用数据文件）
    if results:
        # 获取脚本所在目录的绝对路径
        script_dir = os.path.dirname(os.path.abspath(__file__))
        # data目录在项目根目录下（scripts的父目录）
        data_dir = os.path.join(os.path.dirname(script_dir), 'data')
        # 覆盖通用数据文件，添加锁定标记
        locked_file = os.path.join(data_dir, 'market_indices_scenario_data_v2.json')

        with open(locked_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)

        print()
        print(f"✅ 已保存锁定指数数据到: {locked_file}")
        print(f"   发行日: {issue_date}")
        print(f"   数据截止日: {data_end_date}")
        print(f"   包含指数数量: {len(results)}")
        print(f"   ⚠️ 注意：已覆盖通用数据文件，所有分析将使用锁定数据")

        return results
    else:
        print("❌ 没有成功计算任何指数数据")
        return None


def rebuild_indices_data():
    """
    重新计算所有指数数据（从Tushare API获取）
    """
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
    print("🔄 重新计算市场指数数据（从Tushare API获取）")
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

            # 波动率（所有窗口期）
            vol_20d = calculate_volatility(df, 20)
            vol_60d = calculate_volatility(df, 60)
            vol_120d = calculate_volatility(df, 120)
            vol_250d = calculate_volatility(df, 250)

            # 年化收益率和区间对数收益率（所有窗口期）
            ann_20d, period_20d = calculate_annual_return_v2(df, 20)
            ann_60d, period_60d = calculate_annual_return_v2(df, 60)
            ann_120d, period_120d = calculate_annual_return_v2(df, 120)
            ann_250d, period_250d = calculate_annual_return_v2(df, 250)

            # 计算MA
            ma_20 = df['close'].iloc[-20:].mean()
            ma_60 = df['close'].iloc[-60:].mean()
            ma_120 = df['close'].iloc[-120:].mean()
            ma_250 = df['close'].iloc[-250:].mean()

            # 计算胜率
            recent_20 = df['close'].iloc[-20:]
            win_days_20 = (recent_20.diff() > 0).sum()
            win_rate_20d = win_days_20 / 20 if len(recent_20) >= 2 else 0.5

            recent_60 = df['close'].iloc[-60:]
            win_days_60 = (recent_60.diff() > 0).sum()
            win_rate_60d = win_days_60 / 60 if len(recent_60) >= 2 else 0.5

            recent_120 = df['close'].iloc[-120:]
            win_days_120 = (recent_120.diff() > 0).sum()
            win_rate_120d = win_days_120 / 120 if len(recent_120) >= 2 else 0.5

            recent_250 = df['close'].iloc[-250:]
            win_days_250 = (recent_250.diff() > 0).sum()
            win_rate_250d = win_days_250 / 250 if len(recent_250) >= 2 else 0.5

            # 构建完整结果（包含所有窗口期）
            results[name] = {
                'current_level': current_close,
                # 波动率
                'volatility_20d': vol_20d if not np.isnan(vol_20d) else 0.30,
                'volatility_60d': vol_60d if not np.isnan(vol_60d) else 0.30,
                'volatility_120d': vol_120d if not np.isnan(vol_120d) else 0.35,
                'volatility_250d': vol_250d if not np.isnan(vol_250d) else 0.35,
                # 年化对数收益率
                'return_20d': ann_20d if not np.isnan(ann_20d) else 0,
                'return_60d': ann_60d if not np.isnan(ann_60d) else 0,
                'return_120d': ann_120d if not np.isnan(ann_120d) else 0,
                'return_250d': ann_250d if not np.isnan(ann_250d) else 0,
                # 区间对数收益率（新增字段）
                'period_log_return_20d': period_20d if not np.isnan(period_20d) else 0,
                'period_log_return_60d': period_60d if not np.isnan(period_60d) else 0,
                'period_log_return_120d': period_120d if not np.isnan(period_120d) else 0,
                'period_log_return_250d': period_250d if not np.isnan(period_250d) else 0,
                # 移动平均线
                'ma_20': ma_20,
                'ma_60': ma_60,
                'ma_120': ma_120,
                'ma_250': ma_250,
                # 胜率
                'win_rate_20d': win_rate_20d,
                'win_rate_60d': win_rate_60d,
                'win_rate_120d': win_rate_120d,
                'win_rate_250d': win_rate_250d,
            }

            print(f"  ✅ 完成: 20日收益率={ann_20d*100:+.2f}%, 250日收益率={ann_250d*100:+.2f}%")

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
        print("📊 数据摘要（250日年化收益率）：")
        print(f"{'指数':<12} {'当前点位':<12} {'波动率':<12} {'收益率':<12}")
        print('-'*70)
        for name, data in results.items():
            print(f"{name:<12} {data['current_level']:>10.2f}    {data['volatility_250d']*100:>10.2f}%    {data['return_250d']*100:>+10.2f}%")

        print()
        print("📋 包含窗口期：20日、60日、120日、250日（所有指标）")
        print("📋 包含指标：波动率、收益率、移动平均线、胜率")
        print("="*70)
    else:
        print("❌ 没有成功计算任何指数数据")

    return results


def update_existing_indices_data():
    """
    更新现有的指数数据文件，补充缺失的窗口期指标
    """
    # 获取脚本所在目录的绝对路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # data目录在项目根目录下（scripts的父目录）
    data_dir = os.path.join(os.path.dirname(script_dir), 'data')
    data_file = os.path.join(data_dir, 'market_indices_scenario_data.json')

    if not os.path.exists(data_file):
        print(f"❌ 文件不存在: {data_file}")
        print(f"   提示: 请先运行 --mode rebuild 生成数据文件")
        return

    # 读取现有数据
    with open(data_file, 'r', encoding='utf-8') as f:
        indices_data = json.load(f)

    print(f"✅ 已加载 {len(indices_data)} 个指数数据")
    print()

    # 更新每个指数的数据
    for index_name, data in indices_data.items():
        updated = False

        # 添加/更新20日指标（如果缺失）
        if 'volatility_20d' not in data:
            if 'volatility_30d' in data:
                data['volatility_20d'] = data['volatility_30d']
                updated = True
            else:
                data['volatility_20d'] = 0.30  # 默认值
                updated = True

        if 'return_20d' not in data:
            if 'return_30d' in data:
                data['return_20d'] = data['return_30d']
                updated = True
            else:
                data['return_20d'] = 0
                updated = True

        if 'ma_20' not in data:
            if 'ma_30' in data:
                data['ma_20'] = data['ma_30']
                updated = True
            else:
                data['ma_20'] = data.get('current_level', 3000)
                updated = True

        # 添加/更新250日指标（如果缺失）
        if 'volatility_250d' not in data:
            if 'volatility_180d' in data:
                data['volatility_250d'] = data['volatility_180d']
                updated = True
            else:
                data['volatility_250d'] = 0.35  # 默认值
                updated = True

        if 'return_250d' not in data:
            if 'return_180d' in data:
                data['return_250d'] = data['return_180d']
                updated = True
            else:
                data['return_250d'] = 0
                updated = True

        if 'ma_250' not in data:
            if 'ma_180' in data:
                data['ma_250'] = data['ma_180']
                updated = True
            else:
                data['ma_250'] = data.get('current_level', 3000)
                updated = True

        # 添加win_rate_20d和win_rate_250d（如果缺失）
        if 'win_rate_20d' not in data:
            if 'win_rate_60d' in data:
                data['win_rate_20d'] = data['win_rate_60d']
                updated = True
            else:
                data['win_rate_20d'] = 0.50  # 默认50%
                updated = True

        if 'win_rate_250d' not in data:
            if 'win_rate_60d' in data:
                data['win_rate_250d'] = data['win_rate_60d']
                updated = True
            else:
                data['win_rate_250d'] = 0.50  # 默认50%
                updated = True

        if updated:
            print(f"  {index_name}: ✅ 已补充缺失指标")
        else:
            print(f"  {index_name}: ✓ 已完整")

    # 保存更新后的数据
    with open(data_file, 'w', encoding='utf-8') as f:
        json.dump(indices_data, f, indent=2, ensure_ascii=False)

    print()
    print("="*70)
    print(f"✅ 已更新指数数据文件: {data_file}")
    print()
    print("📋 更新后的数据结构包含:")
    print("   - volatility_20d, volatility_60d, volatility_120d, volatility_250d")
    print("   - return_20d, return_60d, return_120d, return_250d")
    print("   - ma_20, ma_60, ma_120, ma_250")
    print("   - win_rate_20d, win_rate_60d, win_rate_120d, win_rate_250d")
    print("="*70)


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='市场指数数据更新脚本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用示例:
  python update_indices_data.py                           # 默认：重新获取最新数据
  python update_indices_data.py --mode rebuild            # 重新获取最新数据
  python update_indices_data.py --mode update             # 更新现有文件
  python update_indices_data.py --mode locked --stock 300735.SZ --issue 20260407  # 生成锁定数据
        '''
    )

    parser.add_argument(
        '--mode',
        choices=['rebuild', 'update', 'locked'],
        default='rebuild',
        help='运行模式: rebuild(重新获取最新数据) 或 update(更新现有文件) 或 locked(生成基于发行日的锁定数据)'
    )

    parser.add_argument(
        '--stock',
        type=str,
        help='股票代码（锁定模式必需，如: 300735.SZ）'
    )

    parser.add_argument(
        '--issue',
        type=str,
        help='发行日期（锁定模式必需，格式: YYYYMMDD，如: 20260407）'
    )

    args = parser.parse_args()

    print("市场指数数据更新工具")
    print("="*70)
    print(f"运行模式: {args.mode}")
    print()

    if args.mode == 'locked':
        # 锁定模式：生成基于发行日的锁定指数数据
        if not args.stock or not args.issue:
            print("❌ 锁定模式需要 --stock 和 --issue 参数")
            print("   示例: python update_indices_data.py --mode locked --stock 300735.SZ --issue 20260407")
            return
        rebuild_locked_indices_data(args.stock, args.issue)
    elif args.mode == 'rebuild':
        rebuild_indices_data()
    else:
        update_existing_indices_data()


if __name__ == '__main__':
    main()
