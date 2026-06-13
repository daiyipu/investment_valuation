"""波二信号函数(纯函数, 无 IO)。

closes = qfq 收盘序列(oldest→newest, 末项=当日 date)。
逻辑: 经典波浪——一浪上涨(≥20%) → 二浪回调(斐波那契0.382~0.618) → 三浪启动(突破回调高点)。
长周期须上行(close>MA120 且 MA120 上行)。
"""
import numpy as np

from .indicators import sma, swing_high_low, fib_retracement
from .params import WAVE2 as P


def wave2_signal(closes, p=P):
    """closes: qfq 收盘序列(oldest→newest, 末项=当日)。

    返回 {trigger, score, gain, retr, breakout}。
    """
    c = np.asarray(closes, dtype=float)
    n = len(c)
    res = dict(trigger=False, score=0.0, gain=0.0, retr=0.0, breakout='')
    if n < p['ma_long2'] + 5:
        return res
    ma_long = sma(c, p['ma_long']).values
    ma_mid = sma(c, p['ma_mid']).values
    ma_fast = sma(c, p['ma_fast']).values
    ma_slow = sma(c, p['ma_slow']).values
    today = n - 1

    # 1. 长周期上行: close>MA120 且 MA120 近20日斜率>0
    if not (c[today] > ma_long[today] and ma_long[today] > ma_long[today - 20]):
        return res

    # 2. 一浪: [date-120, date-30] 区间低→高
    lo_i = max(0, today - p['lookback_start'])
    hi_i = max(lo_i + 1, today - p['lookback_end'])
    sh = swing_high_low(c, lo_i, hi_i)
    if sh is None:
        return res
    w1_high, w1_high_pos, w1_low, w1_low_pos = sh
    gain = (w1_high - w1_low) / w1_low if w1_low > 0 else 0.0
    res['gain'] = gain
    if gain < p['wave1_min_gain']:
        return res
    # 要求 high 在 low 之后(先低后高=上涨浪)
    if w1_high_pos <= w1_low_pos:
        return res

    # 3. 二浪回撤: 从 w1_high 到今日前的最低; 回撤 0.382~0.618; 不破 w1_low
    seg = c[w1_high_pos:today]  # 含 high, 不含 today
    if len(seg) == 0:
        return res
    pull_low = float(seg.min())
    retr = fib_retracement(w1_low, w1_high, pull_low)
    res['retr'] = retr
    if not (p['retr_lo'] <= retr <= p['retr_hi']):
        return res
    if pull_low <= w1_low:        # 铁律: 不破一浪起点
        return res
    # 回调期守 MA60(辅助)
    pull_low_pos = w1_high_pos + int(np.argmin(seg))
    if not np.isnan(ma_mid[pull_low_pos]) and c[pull_low_pos] < ma_mid[pull_low_pos] * 0.98:
        # 允许略微破 MA60 但不破 w1_low; 这里仅作软约束, 可关
        pass

    # 4. 三浪启动(当日): close 突破回调期高点(自 w1_high 以来最高, 不含 today) OR MA5 上穿 MA20
    pull_high = float(seg.max())  # 回调期最高(通常=w1_high 附近)
    breakout_price = c[today] > pull_high * 0.999  # 突破回调期高点
    breakout_ma = (ma_fast[today] > ma_slow[today]) and (ma_fast[today - 1] <= ma_slow[today - 1])
    if breakout_price:
        res['breakout'] = 'price'
    elif breakout_ma and p['require_ma_confirm']:
        res['breakout'] = 'ma_cross'
    else:
        return res

    res['trigger'] = True
    res['score'] = round(gain * (1.0 if res['breakout'] == 'price' else 0.7), 4)
    return res
