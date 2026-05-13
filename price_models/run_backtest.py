#!/usr/bin/env python3
"""
ARIMA-LSTM 混合模型回测脚本

使用方法:
  # 单股回测（对比三种模型）
  python run_backtest.py --stock 300380.SZ --name 安硕信息

  # 批量回测（DB中所有股票）
  python run_backtest.py --all

  # 指定模型
  python run_backtest.py --stock 300735.SZ --models arima,hybrid

  # 滚动验证
  python run_backtest.py --stock 300380.SZ --rolling
"""

import sys
import os
import argparse
import numpy as np

# 添加当前目录到路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)


def run_single(stock_code, stock_name='', models=None, rolling=False,
               window=250, horizon=120, step=30, verbose=True):
    """单股回测"""
    from data_loader import load_price_series, load_log_returns
    from models.arima_model import fit_arima, forecast_arima, STATSMODELS_AVAILABLE
    from models.lstm_model import forecast_pure_lstm
    from models.hybrid_model import forecast_hybrid, TORCH_AVAILABLE
    from evaluation.metrics import compare_models, format_comparison

    if models is None:
        models = ['arima', 'lstm', 'hybrid']

    print(f"\n{'='*60}")
    print(f"  股票: {stock_code} {stock_name}")
    print(f"  模型: {', '.join(models)}")
    print(f"  PyTorch: {'可用' if TORCH_AVAILABLE else '不可用'}")
    print(f"  statsmodels: {'可用' if STATSMODELS_AVAILABLE else '不可用'}")
    print(f"{'='*60}")

    # 加载数据
    if rolling:
        prices = load_price_series(stock_code, window=None)  # 滚动验证需要全量数据
        if prices is None:
            print("❌ 无法加载价格数据")
            return
        print(f"  加载全量价格数据: {len(prices)} 个交易日")

        from evaluation.rolling_validator import run_rolling_validation, format_summary
        result = run_rolling_validation(
            prices, window=window, horizon=horizon, step=step,
            models=models, verbose=verbose
        )
        print(format_summary(result['summary']))
        return result

    else:
        log_returns = load_log_returns(stock_code, window=window)
        if log_returns is None:
            print("❌ 无法加载收益率数据")
            return

        print(f"  数据窗口: {len(log_returns)} 个交易日")
        recent_120d = log_returns.iloc[-120:].mean() * 252
        print(f"  近120日实际漂移率: {recent_120d*100:+.2f}%")
        print()

        results = {}

        # ARIMA
        if 'arima' in models and STATSMODELS_AVAILABLE:
            try:
                arima_fit = fit_arima(log_returns, auto_find=True)
                arima_fc = forecast_arima(arima_fit['fitted'], horizon=horizon)
                results['arima'] = arima_fc['annualized_drift']
                print(f"  ARIMA {arima_fit['order']}: {arima_fc['annualized_drift']*100:+.2f}%")
            except Exception as e:
                print(f"  ARIMA: 失败 ({e})")

        # 纯LSTM
        if 'lstm' in models:
            lstm_result = forecast_pure_lstm(log_returns, horizon=horizon, verbose=verbose)
            results['lstm'] = lstm_result['annualized_drift']
            status = "✓" if lstm_result['model_fitted'] else f"降级({lstm_result['error']})"
            print(f"  纯LSTM: {lstm_result['annualized_drift']*100:+.2f}% ({status})")

        # 混合模型
        if 'hybrid' in models:
            hybrid_result = forecast_hybrid(log_returns, horizon=horizon, verbose=verbose)
            results['hybrid'] = hybrid_result['annualized_drift']
            blend = hybrid_result.get('blend_weight', 0)
            err = hybrid_result.get('error', '')
            if err:
                print(f"  ARIMA-LSTM: {hybrid_result['annualized_drift']*100:+.2f}% (降级: {err})")
            else:
                print(f"  ARIMA-LSTM: {hybrid_result['annualized_drift']*100:+.2f}% "
                      f"(ARIMA={hybrid_result['arima_drift']*100:+.2f}% + "
                      f"LSTM修正={hybrid_result['lstm_correction']*100:+.2f}%, "
                      f"blend={blend:.2f})")
                if hybrid_result.get('trend_divergence'):
                    print(f"  ⚠️ 预测方向与近120日趋势({recent_120d*100:+.2f}%)矛盾")

        # 对比
        if len(results) > 1:
            comparison = compare_models(results, recent_120d)
            print(format_comparison(comparison, title=f"{stock_code} 三模型对比（基准=近120日实际）"))

        return results


def run_batch(models=None, window=250):
    """批量回测"""
    from data_loader import list_stocks_with_data, load_log_returns
    from models.hybrid_model import forecast_hybrid

    stocks = list_stocks_with_data()
    if not stocks:
        print("❌ DB中无股票数据")
        return

    print(f"\n共 {len(stocks)} 只股票，开始批量回测...")
    results = []

    for stock_code, stock_name in stocks:
        name = stock_name or stock_code
        print(f"\n--- {stock_code} {name} ---")
        try:
            lr = load_log_returns(stock_code, window=window)
            if lr is None:
                continue

            recent = lr.iloc[-120:].mean() * 252
            hybrid = forecast_hybrid(lr, horizon=120, verbose=False)
            arima_d = hybrid.get('arima_drift', 0)
            hybrid_d = hybrid['annualized_drift']
            blend = hybrid.get('blend_weight', 0)
            divergence = hybrid.get('trend_divergence', False)

            results.append({
                'stock_code': stock_code,
                'stock_name': name,
                'recent_120d': recent,
                'arima': arima_d,
                'hybrid': hybrid_d,
                'lstm_correction': hybrid.get('lstm_correction', 0),
                'blend_weight': blend,
                'divergence': divergence,
                'error': hybrid.get('error'),
            })

            flag = "⚠️方向矛盾" if divergence else ""
            err = f" ({hybrid['error']})" if hybrid.get('error') else ""
            print(f"  近120日: {recent*100:+.2f}% | ARIMA: {arima_d*100:+.2f}% | "
                  f"混合: {hybrid_d*100:+.2f}% | blend={blend:.2f}{flag}{err}")

        except Exception as e:
            print(f"  ❌ 失败: {e}")

    # 汇总
    if results:
        valid = [r for r in results if r['error'] is None]
        if valid:
            arima_correct = sum(1 for r in valid
                                if np.sign(r['arima']) == np.sign(r['recent_120d']))
            hybrid_correct = sum(1 for r in valid
                                 if np.sign(r['hybrid']) == np.sign(r['recent_120d']))
            print(f"\n{'='*60}")
            print(f"  批量回测汇总 ({len(valid)}/{len(results)} 成功)")
            print(f"  ARIMA方向准确率: {arima_correct}/{len(valid)} = {arima_correct/len(valid)*100:.1f}%")
            print(f"  混合方向准确率:  {hybrid_correct}/{len(valid)} = {hybrid_correct/len(valid)*100:.1f}%")
            print(f"{'='*60}")


def main():
    parser = argparse.ArgumentParser(description='ARIMA-LSTM 混合模型回测')
    parser.add_argument('--stock', type=str, help='股票代码（如 300380.SZ）')
    parser.add_argument('--name', type=str, default='', help='股票名称')
    parser.add_argument('--all', action='store_true', help='批量回测所有股票')
    parser.add_argument('--models', type=str, default='arima,lstm,hybrid',
                        help='模型列表（逗号分隔，默认 arima,lstm,hybrid）')
    parser.add_argument('--rolling', action='store_true', help='滚动验证模式')
    parser.add_argument('--window', type=int, default=250, help='训练窗口（默认250）')
    parser.add_argument('--horizon', type=int, default=120, help='预测期数（默认120）')
    parser.add_argument('--step', type=int, default=30, help='滚动步长（默认30）')
    parser.add_argument('-q', '--quiet', action='store_true', help='安静模式')

    args = parser.parse_args()
    models = [m.strip() for m in args.models.split(',')]

    if args.all:
        run_batch(models=models, window=args.window)
    elif args.stock:
        run_single(
            args.stock, args.name, models=models,
            rolling=args.rolling, window=args.window,
            horizon=args.horizon, step=args.step,
            verbose=not args.quiet
        )
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
