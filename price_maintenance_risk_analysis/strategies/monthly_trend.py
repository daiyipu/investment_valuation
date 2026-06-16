"""月线 10月均线趋势信号(纯函数, 无 IO)。

monthly_closes = hfq 月度收盘序列(oldest→newest, 末项=报价日最近完整月)。
趋势 = MA10 的 3 月净变化(slope3%) + ±delta 滞后带防单月跳动; 过阈上行=1。
中长期 regime 判断(标签 regime 驱动, 个股自身月线趋势刻画其 regime, 故预期 IV 优于日线级择时)。
"""
import numpy as np
import pandas as pd


def trend10m(closes, delta=0.005):
    """closes: 月度收盘(oldest→newest, hfq)。

    返回 {ma10, slope3_pct, trend_up}:
      ma10       = 末月 10月均线
      slope3_pct = (MA10[t] - MA10[t-3]) / MA10[t-3] × 100  (3月净变化, 百分数, 与 return_7m 同口径)
      trend_up   = 0/1, MA10 经 ±delta 滞过带后的上行状态(防单月跳动; delta 为小数阈值)
    不足 13 月(算不出 MA10+前3月) → ma10/slope3 为 NaN, trend_up=0。
    """
    c = np.asarray(closes, dtype=float)
    res = dict(ma10=np.nan, slope3_pct=np.nan, trend_up=0)
    if len(c) < 13:                       # MA10(需10月) + slope3(再需前3月)
        return res
    ma10 = pd.Series(c).rolling(10).mean()
    today = len(c) - 1
    if pd.isna(ma10.iloc[today]) or pd.isna(ma10.iloc[today - 3]):
        return res
    res['ma10'] = float(ma10.iloc[today])
    res['slope3_pct'] = float((ma10.iloc[today] - ma10.iloc[today - 3]) / ma10.iloc[today - 3] * 100)
    # 滞后带状态机: slope3% > +delta 转 up, < -delta 转 down, 区间内维持上一状态(防跳动)
    state = 0
    for v in ma10.pct_change(3):
        if np.isnan(v):
            continue
        if v > delta:
            state = 1
        elif v < -delta:
            state = 0
    res['trend_up'] = state
    return res
