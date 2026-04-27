#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
第七章 - 蒙特卡洛模拟

包含:
- 7.1 GBM股价模拟
- 7.2 ARIMA预测
- 7.3 GARCH波动率预测
- 7.4 多时间窗口MC: 20/60/120/250日
- 7.5 MC路径图
- 7.6 溢价分布模拟
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from module_utils import (
    add_title, add_paragraph, add_table_data, add_image, add_section_break,
    generate_mc_simulation_chart,
)


def _gbm_simulate(S0, mu, sigma, n_days, n_sims, seed=42):
    """几何布朗运动模拟。

    Args:
        S0: 初始价格
        mu: 年化漂移率
        sigma: 年化波动率
        n_days: 交易日数
        n_sims: 模拟次数
        seed: 随机种子

    Returns:
        paths: ndarray (n_days+1, n_sims)
    """
    dt = 1.0 / 252
    np.random.seed(seed)
    z = np.random.standard_normal((n_days, n_sims))
    log_returns = (mu - 0.5 * sigma ** 2) * dt + sigma * np.sqrt(dt) * z
    log_prices = np.zeros((n_days + 1, n_sims))
    log_prices[0] = np.log(S0)
    log_prices[1:] = np.log(S0) + np.cumsum(log_returns, axis=0)
    return np.exp(log_prices)


def generate_chapter(context):
    """生成第七章：蒙特卡洛模拟

    Args:
        context: dict with keys as specified in project main.py

    Returns:
        context (updated, results['mc_result'] populated)
    """
    document = context['document']
    font_prop = context['font_prop']
    config = context.get('config', {})
    price_data = context.get('price_data')
    daily_basic = context.get('daily_basic', {})
    IMAGES_DIR = context['IMAGES_DIR']
    results = context.setdefault('results', {})

    stock_name = context['stock_name']
    stock_code = context['stock_code']

    mc_cfg = config.get('monte_carlo', {})
    n_simulations = mc_cfg.get('n_simulations', 10000)
    time_windows = mc_cfg.get('time_windows', [20, 60, 120, 250])

    # ==================== 七、蒙特卡洛模拟 ====================
    add_title(document, '七、蒙特卡洛模拟', level=1)
    add_paragraph(document, (
        '本章节使用蒙特卡洛方法，基于几何布朗运动（GBM）模型模拟股价路径，'
        '并结合ARIMA漂移率预测和GARCH波动率预测，评估未来股价分布和投资风险收益。'
    ))

    # 获取当前价格和收益率参数
    current_price = daily_basic.get('close', None)
    if current_price is None and price_data is not None and not price_data.empty:
        price_data_sorted = price_data.sort_values('trade_date', ascending=False)
        current_price = float(price_data_sorted.iloc[0]['close'])

    # 从price_data计算历史mu和sigma
    mu = 0.0
    sigma = 0.30

    if price_data is not None and not price_data.empty:
        prices = price_data.sort_values('trade_date')['close'].values
        if len(prices) > 20:
            log_returns = np.diff(np.log(prices.astype(float)))
            mu = np.mean(log_returns) * 252
            sigma = np.std(log_returns) * np.sqrt(252)

    add_paragraph(document, f'当前价格: {current_price:.2f}元' if current_price else '当前价格: 不可用')
    add_paragraph(document, f'历史年化漂移率(mu): {mu * 100:.2f}%')
    add_paragraph(document, f'历史年化波动率(sigma): {sigma * 100:.2f}%')
    add_paragraph(document, f'模拟次数: {n_simulations:,}')

    # ------------------------------------------------------------------
    # 7.1  GBM股价模拟
    # ------------------------------------------------------------------
    add_title(document, '7.1 GBM股价模拟', level=2)
    add_paragraph(document, (
        '基于几何布朗运动模型 dS = mu*S*dt + sigma*S*dW，'
        f'模拟{n_simulations:,}条120个交易日的股价路径。'
    ))

    mc_paths = None
    final_prices = None

    if current_price and current_price > 0:
        mc_paths = _gbm_simulate(current_price, mu, sigma, 120, n_simulations, seed=42)
        final_prices = mc_paths[-1, :]

        # 统计
        mean_price = np.mean(final_prices)
        median_price = np.median(final_prices)
        p5 = np.percentile(final_prices, 5)
        p25 = np.percentile(final_prices, 25)
        p75 = np.percentile(final_prices, 75)
        p95 = np.percentile(final_prices, 95)

        headers = ['统计指标', '数值(元)']
        data = [
            ['均值', f'{mean_price:.2f}'],
            ['中位数', f'{median_price:.2f}'],
            ['5%分位', f'{p5:.2f}'],
            ['25%分位', f'{p25:.2f}'],
            ['75%分位', f'{p75:.2f}'],
            ['95%分位', f'{p95:.2f}'],
            ['当前价', f'{current_price:.2f}'],
        ]
        add_table_data(document, headers, data)
    else:
        add_paragraph(document, '当前价格不可用，跳过GBM模拟。')

    # ------------------------------------------------------------------
    # 7.2  ARIMA预测
    # ------------------------------------------------------------------
    add_title(document, '7.2 ARIMA漂移率预测', level=2)
    add_paragraph(document, (
        '使用ARIMA时间序列模型预测未来120日的年化漂移率。'
    ))

    arima_drift = mu  # fallback
    arima_fitted = False

    try:
        from utils.time_series_forecaster import TimeSeriesForecaster

        if price_data is not None and not price_data.empty:
            prices_series = price_data.sort_values('trade_date')['close'].astype(float)
            prices_series = prices_series.reset_index(drop=True)
            if len(prices_series) >= 100:
                forecaster = TimeSeriesForecaster(prices_series)
                arima_result = forecaster.forecast_drift_with_arima(horizon=120)
                arima_drift = arima_result.get('forecast_drift', mu)
                arima_fitted = arima_result.get('model_fitted', False)

                add_paragraph(document, f'ARIMA预测漂移率（年化）: {arima_drift * 100:.2f}%')
                add_paragraph(document, f'模型拟合状态: {"成功" if arima_fitted else "失败（使用历史均值）"}')
                if arima_fitted and 'aic' in arima_result:
                    add_paragraph(document, f'AIC: {arima_result["aic"]:.2f}')
            else:
                add_paragraph(document, '价格序列数据不足（需至少100个数据点），使用历史均值。')
        else:
            add_paragraph(document, '价格数据不可用，跳过ARIMA预测。')
    except ImportError:
        add_paragraph(document, 'TimeSeriesForecaster模块未安装，使用历史平均漂移率。')
    except Exception as e:
        add_paragraph(document, f'ARIMA预测失败: {e}，使用历史平均漂移率。')

    add_paragraph(document, f'最终使用的漂移率: {arima_drift * 100:.2f}%')

    # ------------------------------------------------------------------
    # 7.3  GARCH波动率预测
    # ------------------------------------------------------------------
    add_title(document, '7.3 GARCH波动率预测', level=2)
    add_paragraph(document, (
        '使用GARCH(1,1)模型预测未来120日的年化波动率。'
    ))

    garch_vol = sigma  # fallback
    garch_fitted = False

    try:
        from utils.time_series_forecaster import TimeSeriesForecaster

        if price_data is not None and not price_data.empty:
            prices_series = price_data.sort_values('trade_date')['close'].astype(float)
            prices_series = prices_series.reset_index(drop=True)
            if len(prices_series) >= 100:
                forecaster_garch = TimeSeriesForecaster(prices_series)
                garch_result = forecaster_garch.forecast_volatility_with_garch(horizon=120)
                garch_vol = garch_result.get('forecast_volatility', sigma)
                garch_fitted = garch_result.get('model_fitted', False)

                add_paragraph(document, f'GARCH预测波动率（年化）: {garch_vol * 100:.2f}%')
                add_paragraph(document, f'模型拟合状态: {"成功" if garch_fitted else "失败（使用历史波动率）"}')
                if garch_fitted:
                    add_paragraph(document, f'模型参数: omega={garch_result.get("omega", 0):.6f}, '
                                           f'alpha={garch_result.get("alpha", 0):.4f}, '
                                           f'beta={garch_result.get("beta", 0):.4f}')
                    persistence = garch_result.get('alpha', 0) + garch_result.get('beta', 0)
                    add_paragraph(document, f'持续性(alpha+beta): {persistence:.4f}')
            else:
                add_paragraph(document, '价格序列数据不足，使用历史波动率。')
        else:
            add_paragraph(document, '价格数据不可用，跳过GARCH预测。')
    except ImportError:
        add_paragraph(document, 'TimeSeriesForecaster模块未安装，使用历史波动率。')
    except Exception as e:
        add_paragraph(document, f'GARCH预测失败: {e}，使用历史波动率。')

    add_paragraph(document, f'最终使用的波动率: {garch_vol * 100:.2f}%')

    # ------------------------------------------------------------------
    # 7.4  多时间窗口MC
    # ------------------------------------------------------------------
    add_title(document, '7.4 多时间窗口蒙特卡洛模拟', level=2)
    add_paragraph(document, (
        '在20/60/120/250日不同时间窗口下运行蒙特卡洛模拟，对比不同持有期的价格分布。'
    ))

    multi_window_results = {}

    if current_price and current_price > 0:
        headers_mw = ['时间窗口', '模拟次数', '均值价格', '中位数价格', '5%分位', '95%分位',
                       'P(涨)>0', f'P(涨)>10%']
        data_mw = []

        for n_days in time_windows:
            paths = _gbm_simulate(current_price, arima_drift, garch_vol, n_days, n_simulations, seed=42)
            fp = paths[-1, :]

            mean_p = np.mean(fp)
            med_p = np.median(fp)
            p5_p = np.percentile(fp, 5)
            p95_p = np.percentile(fp, 95)
            prob_up = float((fp > current_price).mean() * 100)
            prob_up_10 = float((fp > current_price * 1.10).mean() * 100)

            multi_window_results[n_days] = {
                'mean': mean_p,
                'median': med_p,
                'p5': p5_p,
                'p95': p95_p,
                'prob_up': prob_up,
                'prob_up_10': prob_up_10,
                'paths': paths,
                'final_prices': fp,
            }

            data_mw.append([
                f'{n_days}日',
                f'{n_simulations:,}',
                f'{mean_p:.2f}',
                f'{med_p:.2f}',
                f'{p5_p:.2f}',
                f'{p95_p:.2f}',
                f'{prob_up:.1f}%',
                f'{prob_up_10:.1f}%',
            ])

        add_table_data(document, headers_mw, data_mw)

        add_paragraph(document, '')
        add_paragraph(document, (
            f'参数: 漂移率={arima_drift * 100:.2f}%（ARIMA预测）, '
            f'波动率={garch_vol * 100:.2f}%（GARCH预测）'
        ))
    else:
        add_paragraph(document, '当前价格不可用，跳过多时间窗口模拟。')

    # ------------------------------------------------------------------
    # 7.5  MC路径图
    # ------------------------------------------------------------------
    add_title(document, '7.5 蒙特卡洛模拟路径图', level=2)

    mc_chart_path = os.path.join(IMAGES_DIR, '07_mc_simulation_paths.png')

    if mc_paths is not None:
        try:
            generate_mc_simulation_chart(
                mc_paths, mc_chart_path,
                title=f'{stock_name} 蒙特卡洛模拟（120日, {n_simulations:,}次）',
            )
            if os.path.exists(mc_chart_path):
                add_paragraph(document, '图表7.1 蒙特卡洛模拟路径图（120日）')
                add_image(document, mc_chart_path)
        except Exception as e:
            print(f"  生成MC路径图失败: {e}")

        # 额外绘制价格分布直方图
        try:
            fig, ax = plt.subplots(figsize=(10, 5))
            ax.hist(final_prices, bins=80, alpha=0.7, color='steelblue', edgecolor='white', density=True)
            ax.axvline(x=current_price, color='red', linestyle='--', linewidth=2,
                       label=f'当前价 {current_price:.2f}')
            ax.axvline(x=np.mean(final_prices), color='orange', linestyle='--', linewidth=2,
                       label=f'均值 {np.mean(final_prices):.2f}')
            ax.set_xlabel('模拟期末价格(元)', fontproperties=font_prop, fontsize=12)
            ax.set_ylabel('概率密度', fontproperties=font_prop, fontsize=12)
            ax.set_title(f'{stock_name} 120日模拟价格分布', fontproperties=font_prop, fontsize=14, fontweight='bold')
            ax.legend(prop=font_prop, fontsize=10)
            ax.grid(True, alpha=0.3)
            for label in ax.get_xticklabels():
                label.set_fontproperties(font_prop)
            for label in ax.get_yticklabels():
                label.set_fontproperties(font_prop)
            plt.tight_layout()
            dist_path = os.path.join(IMAGES_DIR, '07_mc_price_distribution.png')
            plt.savefig(dist_path, dpi=150, bbox_inches='tight')
            plt.close()
            if os.path.exists(dist_path):
                add_paragraph(document, '图表7.2 120日模拟价格分布')
                add_image(document, dist_path)
        except Exception as e:
            print(f"  生成价格分布图失败: {e}")

    else:
        add_paragraph(document, '模拟路径数据不可用，跳过路径图生成。')

    # ------------------------------------------------------------------
    # 7.6  溢价分布模拟
    # ------------------------------------------------------------------
    add_title(document, '7.6 溢价分布模拟', level=2)
    add_paragraph(document, (
        '计算蒙特卡洛模拟中，股价超过不同目标价格的概率。'
        '目标价格基于当前价格的溢价水平（+5%/+10%/+20%/+30%等）。'
    ))

    if current_price and current_price > 0 and final_prices is not None:
        targets = {
            '当前价格': current_price,
            '溢价5%': current_price * 1.05,
            '溢价10%': current_price * 1.10,
            '溢价20%': current_price * 1.20,
            '溢价30%': current_price * 1.30,
            '折价10%': current_price * 0.90,
            '折价20%': current_price * 0.80,
        }

        headers_pre = ['目标价位', '目标价(元)', 'P(价格>目标)', '概率(%)']
        data_pre = []
        for name, target in targets.items():
            prob = float((final_prices > target).mean())
            data_pre.append([name, f'{target:.2f}', f'{prob:.4f}', f'{prob * 100:.1f}%'])

        add_table_data(document, headers_pre, data_pre)

        # 概率柱状图
        try:
            fig, ax = plt.subplots(figsize=(10, 5))
            x_labels = list(targets.keys())
            probs = [float((final_prices > t).mean()) * 100 for t in targets.values()]
            colors = ['#27ae60' if p > 60 else '#f39c12' if p > 40 else '#e74c3c' for p in probs]
            bars = ax.bar(x_labels, probs, color=colors, alpha=0.8, edgecolor='white')
            ax.axhline(y=50, color='black', linestyle='--', linewidth=1, label='50%概率线')
            ax.set_ylabel('概率(%)', fontproperties=font_prop, fontsize=12)
            ax.set_title(f'{stock_name} 120日溢价概率分布', fontproperties=font_prop, fontsize=14, fontweight='bold')
            ax.legend(prop=font_prop)
            ax.grid(True, alpha=0.3, axis='y')
            for bar, prob_val in zip(bars, probs):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 1,
                        f'{prob_val:.1f}%', ha='center', va='bottom', fontproperties=font_prop, fontsize=9)
            for label in ax.get_xticklabels():
                label.set_fontproperties(font_prop)
                label.set_rotation(15)
            for label in ax.get_yticklabels():
                label.set_fontproperties(font_prop)
            plt.tight_layout()
            premium_path = os.path.join(IMAGES_DIR, '07_premium_distribution.png')
            plt.savefig(premium_path, dpi=150, bbox_inches='tight')
            plt.close()
            if os.path.exists(premium_path):
                add_paragraph(document, '图表7.3 120日溢价概率分布图')
                add_image(document, premium_path)
        except Exception as e:
            print(f"  生成溢价分布图失败: {e}")

        # 关键概率结论
        prob_gt_10 = float((final_prices > current_price * 1.10).mean()) * 100
        prob_lt_10 = float((final_prices < current_price * 0.90).mean()) * 100
        add_paragraph(document, '')
        add_paragraph(document, f'P(120日价格 > 当前价 x 1.10) = {prob_gt_10:.1f}%')
        add_paragraph(document, f'P(120日价格 < 当前价 x 0.90) = {prob_lt_10:.1f}%')
    else:
        add_paragraph(document, '模拟数据不完整，跳过溢价分布分析。')

    add_section_break(document)

    # ==================== 保存结果 ====================
    mc_result = {
        'mu': float(mu),
        'sigma': float(sigma),
        'arima_drift': float(arima_drift),
        'arima_fitted': arima_fitted,
        'garch_vol': float(garch_vol),
        'garch_fitted': garch_fitted,
        'n_simulations': n_simulations,
        'current_price': float(current_price) if current_price else None,
    }

    if final_prices is not None:
        mc_result['mean_price'] = float(np.mean(final_prices))
        mc_result['median_price'] = float(np.median(final_prices))
        mc_result['p5'] = float(np.percentile(final_prices, 5))
        mc_result['p25'] = float(np.percentile(final_prices, 25))
        mc_result['p75'] = float(np.percentile(final_prices, 75))
        mc_result['p95'] = float(np.percentile(final_prices, 95))

    # 多时间窗口结果
    for n_days, mw_data in multi_window_results.items():
        mc_result[f'{n_days}d_mean'] = float(mw_data['mean'])
        mc_result[f'{n_days}d_median'] = float(mw_data['median'])
        mc_result[f'{n_days}d_p5'] = float(mw_data['p5'])
        mc_result[f'{n_days}d_p95'] = float(mw_data['p95'])
        mc_result[f'{n_days}d_prob_up'] = float(mw_data['prob_up'])

    context['results']['mc_result'] = mc_result

    return context
