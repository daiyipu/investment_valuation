"""纯计算工具: MA / 相关性 / swing高低 / 斐波那契回撤。

全部无 IO, 可单测。供 wave2 / resist 信号函数复用。
"""
import numpy as np
import pandas as pd


def sma(values, window):
    """简单移动平均, 返回与 values 等长 Series(前 window-1 为 NaN)。"""
    return pd.Series(values).rolling(window, min_periods=window).mean()


def rolling_corr(a, b, window):
    """a, b 等长序列的 rolling 相关系数。"""
    return pd.Series(a).rolling(window, min_periods=window).corr(pd.Series(b))


def daily_returns(prices):
    """价格序列 → 日收益率(首项 NaN)。"""
    s = pd.Series(prices, dtype=float)
    return s.pct_change()


def swing_high_low(closes, lo_idx, hi_idx):
    """closes[lo_idx:hi_idx] 区间内的最高/最低及其位置。

    返回 (high, high_pos, low, low_pos)；区间空返回 None。
    pos 为原序列绝对下标。
    """
    seg = np.asarray(closes[lo_idx:hi_idx], dtype=float)
    if len(seg) == 0:
        return None
    high_pos = int(np.argmax(seg)) + lo_idx
    low_pos = int(np.argmin(seg)) + lo_idx
    return float(seg.max()), high_pos, float(seg.min()), low_pos


def fib_retracement(wave1_low, wave1_high, current_low):
    """current_low 相对 wave1(低→高)的回撤比例: (high-current)/(high-low)。

    无回撤(回到高点)→0；span<=0→0。
    """
    span = wave1_high - wave1_low
    if span <= 0:
        return 0.0
    return (wave1_high - current_low) / span
