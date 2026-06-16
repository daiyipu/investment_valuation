"""三浪 / 抵抗 策略集中参数。

调阈值只改这里(及对应信号函数内的逻辑), 不动 runner。
【待定】项后续按回测/屏幕结果再校。
"""

# 三浪参数 (【待定】可调)
# 经典波浪: 一浪上涨 → 二浪回调(斐波那契0.382~0.618) → 三浪主升启动
WAVE3 = dict(
    ma_long=120, ma_long2=250, ma_mid=60, ma_fast=5, ma_slow=20,
    lookback_start=120, lookback_end=30,   # 一浪在 [date-120, date-30] 找
    wave1_min_gain=0.20,                    # 一浪涨幅 ≥20%
    retr_lo=0.382, retr_hi=0.618,           # 斐波那契回撤区间
    require_ma_confirm=True,                # 突破辅以 MA5 上穿 MA20
    volume_ratio_min=1.2,                   # 可选量比(有量数据时)
)

# 抵抗参数
# 个股/行业在大盘/行业下跌波段中抗跌 + 相关性脱钩
RESIST = dict(
    baseline_window=60, recent_windows=(5, 10, 20),
    baseline_corr_floor=0.4,                # 基线 corr ≥0.4 才算
    diverge_threshold=0.4,                  # 背离分综合阈值
    drawdown_threshold=-0.05,               # 大盘/行业下跌波段: 回落 >5%
    peak_lookback=30,                       # 近期高点取近 30 日最高
    weight_stock=0.5, weight_sector=0.5,
)
