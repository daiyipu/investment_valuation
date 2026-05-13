"""评估指标 — 方向准确率、MAE、RMSE等"""

import numpy as np


def direction_accuracy(predicted_drift, actual_drift):
    """方向准确率：预测方向与实际方向一致的比例

    Args:
        predicted_drift: 预测的年化漂移率（或列表）
        actual_drift: 实际的年化漂移率（或列表）

    Returns:
        float: 0.0 ~ 1.0
    """
    pred = np.atleast_1d(predicted_drift)
    actual = np.atleast_1d(actual_drift)

    if len(pred) != len(actual):
        raise ValueError("预测和实际长度不一致")

    correct = np.sum(np.sign(pred) == np.sign(actual))
    return correct / len(pred)


def mae(predicted, actual):
    """平均绝对误差"""
    return np.mean(np.abs(np.array(predicted) - np.array(actual)))


def rmse(predicted, actual):
    """均方根误差"""
    return np.sqrt(np.mean((np.array(predicted) - np.array(actual)) ** 2))


def annualized_error(predicted_drift, actual_drift):
    """年化漂移率误差（百分点）"""
    return abs(predicted_drift - actual_drift) * 100


def compare_models(results_dict, actual_drift):
    """对比多种模型的预测结果

    Args:
        results_dict: {'arima': drift, 'lstm': drift, 'hybrid': drift, ...}
        actual_drift: 实际年化漂移率

    Returns:
        dict: 每个模型的评估指标
    """
    comparison = {}
    for name, pred_drift in results_dict.items():
        comparison[name] = {
            'predicted': pred_drift,
            'actual': actual_drift,
            'error_pct': annualized_error(pred_drift, actual_drift),
            'direction_correct': np.sign(pred_drift) == np.sign(actual_drift),
        }
    return comparison


def format_comparison(comparison, title="模型对比"):
    """格式化输出对比结果"""
    lines = [f"\n{'='*60}", f"  {title}", f"{'='*60}"]
    header = f"{'模型':<12} {'预测':>10} {'实际':>10} {'误差(%)':>10} {'方向':>6}"
    lines.append(header)
    lines.append("-" * 60)

    for name, metrics in comparison.items():
        direction = "✓" if metrics['direction_correct'] else "✗"
        lines.append(
            f"{name:<12} {metrics['predicted']*100:>+9.2f}% "
            f"{metrics['actual']*100:>+9.2f}% "
            f"{metrics['error_pct']:>9.2f}  {direction:>4}"
        )

    lines.append(f"{'='*60}")
    return "\n".join(lines)
