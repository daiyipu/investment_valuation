"""波二/抵抗策略信号函数单测(合成序列)。

运行: cd price_maintenance_risk_analysis && python -m pytest tests/test_strategies.py -v
"""
import numpy as np

from strategies.indicators import sma, rolling_corr, fib_retracement, swing_high_low
from strategies.wave2 import wave2_signal
from strategies.resist import resist_score, _diverge_score


# ---------- Task 1: indicators ----------
def test_sma_basic():
    assert list(sma([1, 2, 3, 4], 2).dropna()) == [1.5, 2.5, 3.5]


def test_fib_retracement_half():
    # wave1: 10→20, 回撤到 15 → 回撤 50%
    assert abs(fib_retracement(10, 20, 15) - 0.5) < 1e-9


def test_fib_retracement_no_drop():
    # 没回撤(回到高点) → 0
    assert fib_retracement(10, 20, 20) == 0.0


def test_swing_high_low():
    c = [3, 1, 4, 1, 5, 9, 2, 6]
    h, hp, l, lp = swing_high_low(c, 0, 8)
    assert h == 9 and hp == 5 and l == 1 and lp == 1


def test_rolling_corr_perfect():
    a = [1, 2, 3, 4, 5]; b = [2, 4, 6, 8, 10]
    c = rolling_corr(a, b, 3).dropna()
    assert all(abs(x - 1.0) < 1e-9 for x in c)


# ---------- Task 3: wave2 ----------
def _synth_wave2():
    """构造: 平台→一浪上涨(10→15,+50%)→回调(15→12.5,retr≈0.5)→突破回调高点15。

    lookback 窗口 [today-120, today-30] 排除最后 30 日, 故一浪高点须在 today-30 之前,
    回调落在最后 ~30 日, 今日突破回调期高点。序列长 260(≥ MA250+5 守卫)。
    """
    base = list(np.repeat(10.0, 140))          # 长平台(够 MA120/MA250 守卫)
    ramp = list(np.linspace(10, 15, 51))       # 一浪上涨 +50%(高点 idx190, 早于 today-30)
    pull = list(np.linspace(15, 12.5, 66))     # 二浪回调 retr≈0.5(落在最后区段)
    brk = [13.0, 14.2, 15.2]                   # 反弹+突破回调高点 15
    return base + ramp + pull + brk


def test_wave2_triggers_on_synthetic():
    c = _synth_wave2()
    r = wave2_signal(c)
    assert r['trigger'] is True
    assert r['gain'] >= 0.20
    assert 0.382 <= r['retr'] <= 0.618


def test_wave2_no_trigger_in_downtrend():
    c = list(np.linspace(50, 20, 300))  # 单边下跌
    assert wave2_signal(c)['trigger'] is False


# ---------- Task 4: resist ----------
def test_diverge_score_normalized():
    assert abs(_diverge_score(0.8, 0.4) - 0.5) < 1e-9      # (0.8-0.4)/0.8
    assert _diverge_score(0.3, 0.1) is None                 # 基线<0.4 排除


def test_resist_triggers_when_stock_holds_in_sector_drop():
    rng = np.random.default_rng(0)
    # 前65日: 个股≈行业+小噪声(建立高基线相关); 近20日: 行业渐跌、个股渐涨(背离+抗跌)
    sector_base = rng.normal(0, 0.01, 65)
    stock_base = sector_base + rng.normal(0, 0.002, 65)   # 个股跟随行业 → 基线 corr 高
    market = np.concatenate([rng.normal(0, 0.01, 65), np.full(20, -0.004)])
    sector = np.concatenate([sector_base, np.linspace(-0.001, -0.006, 20)])  # 行业渐跌
    stock = np.concatenate([stock_base, np.linspace(0.0005, 0.002, 20)])     # 个股逆势渐涨
    r = resist_score(stock, sector, market)
    assert r['corr_div_stock'] is not None
    assert r['rel_stock'] is not None and r['rel_stock'] > 0
    assert r['trigger'] is True
