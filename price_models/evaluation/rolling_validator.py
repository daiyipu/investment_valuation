"""滚动验证框架 — 时间序列回测"""

import numpy as np
from models.arima_model import forecast_arima, fit_arima, STATSMODELS_AVAILABLE
from models.lstm_model import forecast_pure_lstm
from models.hybrid_model import forecast_hybrid
from evaluation.metrics import direction_accuracy, mae, format_comparison
from config import HORIZON, SEQ_LENGTH


def run_rolling_validation(prices, window=250, horizon=HORIZON, step=30,
                           models=None, verbose=True):
    """滚动验证

    每次滑动step天，用window天数据训练，预测horizon天。
    用预测区间的实际收益率作为真实值。

    Args:
        prices: 完整价格序列（pd.Series 或 np.array）
        window: 训练窗口大小
        horizon: 预测期数
        step: 滑动步长
        models: 要测试的模型列表，如 ['arima', 'lstm', 'hybrid']
        verbose: 是否打印过程

    Returns:
        dict: {
            'folds': list of fold results,
            'summary': aggregated metrics,
        }
    """
    if models is None:
        models = ['arima', 'lstm', 'hybrid']

    prices = np.array(prices, dtype=np.float64)
    log_returns_all = np.diff(np.log(prices))

    folds = []
    start = 0

    while start + window + horizon <= len(log_returns_all):
        train_returns = log_returns_all[start:start + window]
        # 实际未来horizon天的总收益率
        actual_returns = log_returns_all[start + window:start + window + horizon]
        actual_drift = _annualize(actual_returns)

        fold = {
            'fold': len(folds) + 1,
            'train_start': start,
            'train_end': start + window,
            'test_end': start + window + horizon,
            'actual_drift': actual_drift,
            'predictions': {},
        }

        if verbose:
            print(f"\n--- Fold {fold['fold']} "
                  f"(train:{start}-{start+window}, test:{start+window}-{start+window+horizon}) "
                  f"actual={actual_drift*100:+.2f}% ---")

        # ARIMA
        if 'arima' in models and STATSMODELS_AVAILABLE:
            try:
                arima_fit = fit_arima(train_returns, auto_find=True)
                arima_fc = forecast_arima(arima_fit['fitted'], horizon=horizon)
                fold['predictions']['arima'] = arima_fc['annualized_drift']
                if verbose:
                    print(f"  ARIMA: {arima_fc['annualized_drift']*100:+.2f}%")
            except Exception as e:
                fold['predictions']['arima'] = None
                if verbose:
                    print(f"  ARIMA: 失败 ({e})")

        # 纯LSTM
        if 'lstm' in models:
            try:
                lstm_result = forecast_pure_lstm(
                    train_returns, horizon=horizon, seq_length=SEQ_LENGTH,
                    verbose=False
                )
                fold['predictions']['lstm'] = lstm_result['annualized_drift']
                if verbose:
                    status = "OK" if lstm_result['model_fitted'] else f"fallback({lstm_result['error']})"
                    print(f"  LSTM:  {lstm_result['annualized_drift']*100:+.2f}% ({status})")
            except Exception as e:
                fold['predictions']['lstm'] = None
                if verbose:
                    print(f"  LSTM:  失败 ({e})")

        # 混合模型
        if 'hybrid' in models:
            try:
                hybrid_result = forecast_hybrid(
                    train_returns, horizon=horizon, seq_length=SEQ_LENGTH,
                    verbose=False
                )
                fold['predictions']['hybrid'] = hybrid_result['annualized_drift']
                if verbose:
                    blend = hybrid_result.get('blend_weight', 0)
                    err = hybrid_result.get('error', '')
                    extra = f"blend={blend:.2f}" if not err else f"fallback({err})"
                    print(f"  混合:  {hybrid_result['annualized_drift']*100:+.2f}% ({extra})")
            except Exception as e:
                fold['predictions']['hybrid'] = None
                if verbose:
                    print(f"  混合:  失败 ({e})")

        folds.append(fold)
        start += step

    # 汇总统计
    summary = _aggregate_folds(folds, models)
    return {'folds': folds, 'summary': summary}


def _annualize(returns):
    total_log = returns.sum()
    total_simple = np.exp(total_log) - 1
    if total_simple <= -1:
        return -1.0
    annualized_simple = (1 + total_simple) ** (252 / len(returns)) - 1
    return np.log(1 + annualized_simple)


def _aggregate_folds(folds, models):
    """汇总各模型的统计指标"""
    summary = {}
    actuals = [f['actual_drift'] for f in folds]

    for model_name in models:
        preds = [f['predictions'].get(model_name) for f in folds]
        valid = [(p, a) for p, a in zip(preds, actuals) if p is not None]

        if not valid:
            summary[model_name] = {'n_folds': 0, 'error': 'no valid predictions'}
            continue

        pred_arr = np.array([v[0] for v in valid])
        act_arr = np.array([v[1] for v in valid])

        summary[model_name] = {
            'n_folds': len(valid),
            'direction_accuracy': direction_accuracy(pred_arr, act_arr),
            'mae': mae(pred_arr, act_arr),
            'mean_predicted': pred_arr.mean(),
            'mean_actual': act_arr.mean(),
        }

    return summary


def format_summary(summary):
    """格式化汇总输出"""
    lines = [f"\n{'='*70}", "  滚动验证汇总", f"{'='*70}"]
    lines.append(f"{'模型':<12} {'折数':>4} {'方向准确率':>10} {'MAE':>10} {'平均预测':>12} {'平均实际':>12}")
    lines.append("-" * 70)

    for name, s in summary.items():
        if 'error' in s:
            lines.append(f"{name:<12} {s.get('n_folds', 0):>4}  {s['error']}")
            continue
        lines.append(
            f"{name:<12} {s['n_folds']:>4} {s['direction_accuracy']*100:>9.1f}% "
            f"{s['mae']*100:>9.2f}% {s['mean_predicted']*100:>+11.2f}% "
            f"{s['mean_actual']*100:>+11.2f}%"
        )

    lines.append(f"{'='*70}")
    return "\n".join(lines)
