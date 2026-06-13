# 波二 & 抵抗 选股策略 实施计划

> **For agentic workers:** 实施 task-by-task。波二定义标【待定】，信号函数参数集中放 `strategies/params.py` 便于后续调。

**Goal:** 实现 wave2_signal / resist_score 两个 PIT 信号函数 + 屏幕/特征/预筛三 runner。

**Architecture:** 信号函数为**纯函数**（入参=已取好的价格/收益序列，≤date），数据获取（PIT 取 stock qfq / industry / market）独立在 data_loader。三个 runner 复用同一套信号函数：`screen_strategies.py`(全市场扫描) / `derive_features.py` H&I 类(按报价日回算) / 预筛(屏幕跑标的池)。

**Tech Stack:** Python(venv vnpy, py3.10) / pandas / numpy / MySQL(pymysql) / tushare(pro_bar qfq)。复用 market_data.price_series、industry_daily、market_indices。

---

## File Structure

| 文件 | 职责 |
|---|---|
| `price_maintenance_risk_analysis/strategies/__init__.py` | 包标识 |
| `strategies/params.py` | 所有阈值参数(集中, 便于调) |
| `strategies/indicators.py` | 纯计算工具: MA / corr / swing高低 / 回撤 —— 可单测 |
| `strategies/wave2.py` | `wave2_signal(closes, p) -> dict` 纯函数 |
| `strategies/resist.py` | `resist_score(stock_r, sector_r, market_r, p) -> dict` 纯函数 |
| `strategies/data_loader.py` | PIT 取数: stock qfq收盘 / industry日收益 / market日收益, ≤date 截断 |
| `strategies/screen_strategies.py` | 屏幕 runner: 全市场 or 标的池, 输出 CSV 列表 |
| `ml_training/derive_features.py` | 续编 H(波二)/I(抵抗) 类, 按报价日回算 |
| `tests/test_strategies.py` | 信号函数单测(合成序列) |

设计原则：信号函数纯（无 IO、无 DB），输入序列由 runner 通过 data_loader 提供。这样信号逻辑可单测、三 runner 共享、PIT 由 data_loader 的 ≤date 截断保证。

---

## Task 1: indicators 工具 + 单测

**Files:** Create `strategies/__init__.py`, `strategies/indicators.py`, `tests/test_strategies.py`

`strategies/indicators.py`（纯函数，无 IO）：

```python
import numpy as np
import pandas as pd

def sma(values, window):
    """简单移动平均, 返回与 values 等长 Series(前 window-1 为 NaN)。"""
    return pd.Series(values).rolling(window, min_periods=window).mean()

def rolling_corr(a, b, window):
    """a, b 等长收益/价格序列的 rolling 相关系数。"""
    return pd.Series(a).rolling(window, min_periods=window).corr(pd.Series(b))

def daily_returns(prices):
    """价格序列 → 日收益率(首项 NaN)。"""
    s = pd.Series(prices, dtype=float)
    return s.pct_change()

def swing_high_low(closes, lo_idx, hi_idx):
    """closes[lo_idx:hi_idx] 区间内的最高/最低及其位置。返回 (high, high_pos, low, low_pos)。"""
    seg = closes[lo_idx:hi_idx]
    if len(seg) == 0:
        return None
    high_pos = int(np.argmax(seg)) + lo_idx
    low_pos = int(np.argmin(seg)) + lo_idx
    return float(seg.max()), high_pos, float(seg.min()), low_pos

def fib_retracement(wave1_low, wave1_high, current_low):
    """current_low 相对 wave1(低→高)的回撤比例: (high-current)/(high-low)。"""
    span = wave1_high - wave1_low
    if span <= 0:
        return 0.0
    return (wave1_high - current_low) / span
```

`tests/test_strategies.py`（首版，TDD）：

```python
import numpy as np
from strategies.indicators import sma, rolling_corr, fib_retracement, swing_high_low

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
```

- [ ] 创建三个文件，跑 `cd price_maintenance_risk_analysis && python -m pytest tests/test_strategies.py -v`（vnpy 环境）应全绿。
- [ ] commit: `feat(strategies): indicators 工具+单测`

---

## Task 2: params 集中参数

**Files:** Create `strategies/params.py`

```python
# 波二参数(【待定】可调)
WAVE2 = dict(
    ma_long=120, ma_long2=250, ma_mid=60, ma_fast=5, ma_slow=20,
    lookback_start=120, lookback_end=30,   # 一浪在 [date-120, date-30] 找
    wave1_min_gain=0.20,                    # 一浪涨幅 ≥20%
    retr_lo=0.382, retr_hi=0.618,           # 斐波那契回撤区间
    require_ma_confirm=True,                # 突破辅以 MA5 上穿 MA20
    volume_ratio_min=1.2,                   # 可选量比(有量数据时)
)
# 抵抗参数
RESIST = dict(
    baseline_window=60, recent_windows=(5, 10, 20),
    baseline_corr_floor=0.4,                # 基线 corr ≥0.4 才算
    diverge_threshold=0.4,                  # 背离分综合阈值
    drawdown_threshold=-0.05,               # 大盘/行业下跌波段: 回落 >5%
    peak_lookback=30,                       # 近期高点取近 30 日最高
    weight_stock=0.5, weight_sector=0.5,
)
```

- [ ] 创建文件。commit: `feat(strategies): 集中参数(波二/抵抗, 含【待定】标注)`

---

## Task 3: wave2_signal 信号函数 + 单测

**Files:** Create `strategies/wave2.py`, 追加测试到 `tests/test_strategies.py`

`strategies/wave2.py`（纯函数，closes = qfq 收盘 oldest→newest，最后一个=当日 date）：

```python
import numpy as np
from .indicators import sma, swing_high_low, fib_retracement
from .params import WAVE2 as P

def wave2_signal(closes, p=P):
    """closes: qfq 收盘序列(oldest→newest, 末项=当日)。返回 {trigger, score, gain, retr, breakout}。"""
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
```

测试（合成一段"一浪上涨→回调→突破"序列）：

```python
import numpy as np
from strategies.wave2 import wave2_signal

def _synth_wave2():
    """构造: 平台→一浪上涨(10→14,+40%)→回撤到12.1(约0.475)→突破到14.2。"""
    base = list(np.repeat(10.0, 130))          # 长平台(够 MA120)
    ramp = list(np.linspace(10, 14, 25))       # 一浪上涨 +40%
    pull = list(np.linspace(14, 12.1, 12))     # 回撤 ~47.5%
    brk = [12.1, 13.0, 14.2]                   # 反弹+突破回调高点14
    return base + ramp + pull + brk

def test_wave2_triggers_on_synthetic():
    c = _synth_wave2()
    r = wave2_signal(c)
    assert r['trigger'] is True
    assert r['gain'] >= 0.20
    assert 0.382 <= r['retr'] <= 0.618

def test_wave2_no_trigger_in_downtrend():
    c = list(np.linspace(50, 20, 200))  # 单边下跌
    assert wave2_signal(c)['trigger'] is False
```

- [ ] 创建 wave2.py + 追加测试, 跑 pytest 应绿(合成序列触发、下跌不触发)。
- [ ] commit: `feat(strategies): wave2 信号函数+单测`

> ⚠ 波二定义【待定】：参数全在 params.WAVE2，调整阈值/逻辑只改 wave2.py + params，不动 runner。

---

## Task 4: resist_score 信号函数 + 单测

**Files:** Create `strategies/resist.py`, 追加测试

入参为**日收益率序列**（stock_r, sector_r, market_r，等长，对齐日期，末项=当日）。data_loader 负责取价→算收益对齐。

```python
import numpy as np
import pandas as pd
from .indicators import rolling_corr
from .params import RESIST as P

def _diverge_score(baseline_corr, recent_corr):
    """归一化背离分 = (基线-近W)/基线, 基线<floor 返回 None。"""
    if baseline_corr < P['baseline_corr_floor']:
        return None
    if baseline_corr <= 0:
        return None
    return (baseline_corr - recent_corr) / baseline_corr

def _recent_downswing_cumret(market_r, peak_lookback, thr):
    """找 market 近 peak_lookback 日内最高点 peak, 若 [peak, 末] 累计收益 < thr(如-5%) → 返回该段累计收益; 否则 None(无下跌波段)。"""
    s = pd.Series(market_r, dtype=float)
    n = len(s)
    lo = max(0, n - peak_lookback)
    seg = s.iloc[lo:]
    if len(seg) < 2:
        return None
    peak_pos = int(seg.values.argmax())  # 收益最高的"日"近似高点位置(粗略)
    # 用价格累计更准: 这里用收益率累积近似
    cum = (1 + seg.iloc[peak_pos:]).prod() - 1
    if cum < thr:
        return float(cum)
    return None

def resist_score(stock_r, sector_r, market_r, p=P):
    """返回 {trigger, score, corr_div_stock, corr_div_sector, rel_stock, rel_sector}。"""
    sr = np.asarray(stock_r, dtype=float); kr = np.asarray(sector_r, dtype=float); mr = np.asarray(market_r, dtype=float)
    res = dict(trigger=False, score=0.0, corr_div_stock=None, corr_div_sector=None,
               rel_stock=None, rel_sector=None)
    if len(sr) < p['baseline_window'] + 5:
        return res

    # A. 相关性背离(多窗 5/10/20)
    base_sk = rolling_corr(sr, kr, p['baseline_window']).iloc[-1]
    base_km = rolling_corr(kr, mr, p['baseline_window']).iloc[-1]
    divs_sk, divs_km = [], []
    for w in p['recent_windows']:
        rc_sk = rolling_corr(sr, kr, w).iloc[-1]
        rc_km = rolling_corr(kr, mr, w).iloc[-1]
        d_sk = _diverge_score(base_sk, rc_sk); d_km = _diverge_score(base_km, rc_km)
        if d_sk is not None: divs_sk.append(d_sk)
        if d_km is not None: divs_km.append(d_km)
    div_stock = float(np.mean(divs_sk)) if divs_sk else 0.0
    div_sector = float(np.mean(divs_km)) if divs_km else 0.0
    res['corr_div_stock'] = round(div_stock, 4); res['corr_div_sector'] = round(div_sector, 4)

    # B. 下跌波段抗跌方向
    mkt_down = _recent_downswing_cumret(mr, p['peak_lookback'], p['drawdown_threshold'])
    rel_sector = None; rel_stock = None
    if mkt_down is not None:
        s = pd.Series(mr); n = len(s); lo = max(0, n - p['peak_lookback'])
        seg_sl = slice(lo, n)
        rel_sector = float((1 + pd.Series(kr)[seg_sl]).prod() - 1 - mkt_down)  # 行业-市场(同期累计)
    sec_down = _recent_downswing_cumret(kr, p['peak_lookback'], p['drawdown_threshold'])
    if sec_down is not None:
        s = pd.Series(kr); n = len(s); lo = max(0, n - p['peak_lookback'])
        seg_sl = slice(lo, n)
        rel_stock = float((1 + pd.Series(sr)[seg_sl]).prod() - 1 - sec_down)   # 个股-行业
    res['rel_stock'] = round(rel_stock, 4) if rel_stock is not None else None
    res['rel_sector'] = round(rel_sector, 4) if rel_sector is not None else None

    # C. 综合: 个股背离>阈值 ∧ 个股抗跌>0
    stock_resist = (div_stock * rel_stock) if (rel_stock is not None and rel_stock > 0) else 0.0
    sector_resist = (div_sector * rel_sector) if (rel_sector is not None and rel_sector > 0) else 0.0
    res['trigger'] = (div_stock > p['diverge_threshold']) and (rel_stock is not None and rel_stock > 0)
    res['score'] = round(p['weight_stock'] * stock_resist + p['weight_sector'] * sector_resist, 4)
    return res
```

测试（合成：行业/市场下跌，个股抗跌 + 个股-行业相关性背离）：

```python
import numpy as np
from strategies.resist import resist_score, _diverge_score

def test_diverge_score_normalized():
    assert abs(_diverge_score(0.8, 0.4) - 0.5) < 1e-9      # (0.8-0.4)/0.8
    assert _diverge_score(0.3, 0.1) is None                 # 基线<0.4 排除

def test_resist_triggers_when_stock_holds_in_sector_drop():
    rng = np.random.default_rng(0)
    # 前60日: 个股随行业(高相关); 近20日: 行业跌、个股稳(背离+抗跌)
    n = 85
    market = np.concatenate([rng.normal(0,0.01,65), np.full(20,-0.004)])
    sector = np.concatenate([rng.normal(0,0.01,65), np.full(20,-0.005)])
    stock  = np.concatenate([rng.normal(0,0.01,65), np.full(20, 0.001)])  # 近期逆势小涨
    r = resist_score(stock, sector, market)
    assert r['corr_div_stock'] is not None
    assert r['rel_stock'] is not None and r['rel_stock'] > 0
    assert r['trigger'] is True
```

- [ ] 创建 resist.py + 追加测试, pytest 应绿。
- [ ] commit: `feat(strategies): resist 信号函数+单测`

> 注：`_recent_downswing_cumret` 用收益累积近似下跌波段；若需更准的"高点→当前"价格回撤，data_loader 可直接传价格序列重写（接口已隔离，改 data_loader 不动 resist 主逻辑）。

---

## Task 5: data_loader (PIT 取数)

**Files:** Create `strategies/data_loader.py`

负责按 (stock, date) 取 ≤date 的序列，喂给信号函数。复用现有 market_data.price_series / industry_daily / market_indices。

```python
import json, numpy as np, pandas as pd, pymysql

def _conn():
    return pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')

def stock_qfq_closes(stock_code, date, maxlen=300):
    """个股 ≤date 的 qfq 收盘(oldest→newest)。优先 market_data.price_series, 不足/不含 date 则按股 pro_bar qfq。"""
    # 优先 price_series(market_data, 已 qfq, 报价日锚定——特征侧用)
    conn = _conn(); cur = conn.cursor()
    cur.execute('SELECT price_series FROM market_data WHERE stock_code=%s', (stock_code,))
    row = cur.fetchone(); conn.close()
    if row and row[0]:
        try:
            s = json.loads(row[0])
            return np.array(s[-maxlen:], dtype=float)
        except Exception:
            pass
    # 回退: 实时取 pro_bar qfq(屏幕侧任意 date 用)
    import tushare as ts, os
    ts.set_token(os.environ.get('TUSHARE_TOKEN', 'f2380d8761bcbf165f87b85f04ed105b1bdcf8721574562294671265'))
    df = ts.pro_bar(ts_code=stock_code, end_date=date, adj='qfq', limit=maxlen)
    if df is None or len(df) == 0:
        return np.array([])
    df = df.sort_values('trade_date')
    return df['close'].astype(float).values

def industry_daily_returns(stock_code, date, maxlen=300):
    """该股所属 SW 行业指数 ≤date 的日收益(对齐 trading_date)。"""
    conn = _conn()
    # 取该股行业 index_code
    cur = conn.cursor()
    cur.execute('SELECT sw_l2_code FROM industry_data WHERE stock_code=%s', (stock_code,))
    row = cur.fetchone()
    idx = row[0] if row and row[0] else None
    if not idx:
        conn.close(); return pd.Series(dtype=float)
    df = pd.read_sql('SELECT trade_date, close FROM industry_daily WHERE index_code=%s AND trade_date<=%s ORDER BY trade_date', conn, params=(idx, date))
    conn.close()
    if df.empty: return pd.Series(dtype=float)
    return df['close'].astype(float).pct_change().iloc[-maxlen:]

def market_daily_returns(date, maxlen=300, index_code='000300.SH'):
    """大盘(沪深300) ≤date 日收益。"""
    conn = _conn()
    df = pd.read_sql('SELECT trade_date, close FROM market_indices WHERE index_code=%s AND trade_date<=%s ORDER BY trade_date', conn, params=(index_code, date))
    conn.close()
    if df.empty: return pd.Series(dtype=float)
    return df['close'].astype(float).pct_change().iloc[-maxlen:]
```

> ⚠ 表名/列名（market_indices、industry_daily 的 close/trade_date、index_code 取法）实施时按实际 schema 核对（参考 migrate_to_mysql.py + load_db_features）；上面是骨架，接对真实字段是本 task 的核心动作。

- [ ] 实施时先 `SHOW COLUMNS`/读 migrate_to_mysql.py 核对字段，补对齐逻辑（三个序列须等长、末项对齐 date）。
- [ ] 写 1 个集成测试：取 1 只真实股票+真实 date，断言三个序列非空、等长。
- [ ] commit: `feat(strategies): data_loader PIT 取数`

---

## Task 6: 屏幕 runner

**Files:** Create `strategies/screen_strategies.py`

```python
import argparse, pandas as pd, numpy as np, sys, os
sys.path.insert(0, os.path.dirname(__file__))
from data_loader import stock_qfq_closes, industry_daily_returns, market_daily_returns
from wave2 import wave2_signal
from resist import resist_score

def screen(date, universe=None, out_dir='output'):
    """universe=None → 全市场(从 stocks 表); 否则传股票列表。"""
    import pymysql
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='', database='investment_valuation', charset='utf8mb4')
    codes = universe or pd.read_sql('SELECT stock_code FROM stocks', conn)['stock_code'].tolist()
    conn.close()
    mkt = market_daily_returns(date)
    w2_rows, rs_rows = [], []
    for code in codes:
        try:
            c = stock_qfq_closes(code, date)
            if len(c) > 250:
                r = wave2_signal(c)
                if r['trigger']: w2_rows.append({'股票代码':code, **r})
            ind = industry_daily_returns(code, date)
            if len(ind) > 60 and len(mkt) > 60:
                n = min(len(ind), len(mkt))
                r = resist_score(ind.iloc[-n:].values, ind.iloc[-n:].values, mkt.iloc[-n:].values)  # 注:stock_r 暂用 ind 占位, 实施时换 stock_qfq→收益
                if r['trigger']: rs_rows.append({'股票代码':code, **r})
        except Exception as e:
            print(f'{code} skip: {e}')
    pd.DataFrame(w2_rows).to_csv(f'{out_dir}/wave2_list_{date}.csv', index=False, encoding='utf-8-sig')
    pd.DataFrame(rs_rows).to_csv(f'{out_dir}/resist_list_{date}.csv', index=False, encoding='utf-8-sig')
    print(f'波二 {len(w2_rows)} / 抵抗 {len(rs_rows)} 只, 已输出 CSV')

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--date', required=True)
    ap.add_argument('--universe', default=None, help='逗号分隔股票代码, 缺省全市场')
    args = ap.parse_args()
    uni = args.universe.split(',') if args.universe else None
    screen(args.date, uni)
```

> ⚠ 占位需修：resist_score 的 stock_r 应是**个股**日收益(stock_qfq→pct_change)，不是 ind。实施时补 `stock_r = pd.Series(c).pct_change()` 并与 ind/mkt 末段对齐。这是本 task 必须修正的点（骨架为简洁暂用 ind 占位）。

- [ ] 修正 stock_r 取个股收益、三序列对齐。
- [ ] 跑一个历史 date（如 20240601）+ 小标的池，人工抽查输出列表合理性。
- [ ] commit: `feat(strategies): 屏幕 runner(波二+抵抗列表)`

---

## Task 7: 特征接入 (derive_features H&I 类)

**Files:** Modify `ml_training/derive_features.py`（末尾续编，仿 F/G 类）

在 derive 主流程末尾（现有 derive_market_index_features 之后），按**每个样本报价日**回算波二/抵抗，写回 df：

```python
def derive_strategy_signals(df):
    """按每行报价日 PIT 回算 波二/抵抗 信号(特征侧)。复用 strategies 信号函数 + 现有价格序列列。"""
    import sys, os
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'strategies'))
    from wave2 import wave2_signal
    from resist import resist_score
    print('\n  H/I类: 波二/抵抗策略信号...')
    # 价格序列来源: df 里若有 market_data 的 MA/价格列可用; 否则按报价日取 price_series
    # 这里给骨架: 遍历行, 取该股 ≤报价日 qfq 收盘 + 行业/大盘收益, 调信号函数
    w2t, w2s = [], []
    for _, row in df.iterrows():
        code = str(row.get('股票代码', '')); date = str(row.get('报价日', ''))
        # 取 closes / industry_r / market_r (同 data_loader, 按报价日)
        # closes = ...; w2 = wave2_signal(closes)
        # ... 调用并收集
        pass  # 实施时填: 取数(复用 data_loader) + 调 wave2_signal / resist_score
    # df['波二_trigger']=...; df['波二_score']=...; df['抵抗_综合']=...
    return df
```

> ⚠ 此 task 是接入骨架——实施时把 pass 段填实（复用 data_loader 取数 + 调信号函数 + 写 5 个特征列），并在 `derive_features.main` 的 try 块追加 `df = derive_strategy_signals(df)`。覆盖率低的特征按手动清单原则进 feature_exclusions。

- [ ] 填实 derive_strategy_signals（取数+调信号+写列）+ 接入 main。
- [ ] 跑 derive_features 一次，看 5 个新特征覆盖率/非空率。
- [ ] commit: `feat(derive): 接入波二/抵抗策略信号(H/I类)`

---

## Task 8: 端到端验证

- [ ] **屏幕验证**：`python strategies/screen_strategies.py --date <最近收盘日>`，看波二/抵抗列表，人工抽查 3-5 只（熟悉的股票是否符合"二浪末端突破"/"抗跌脱钩"）。
- [ ] **特征验证**：重建 features_derived（含 H/I），跑 `compare_selection.py`/`train_scorecard.py`，看波二/抵抗特征 IV、是否入选评分卡、validate AUC 是否提升。
- [ ] **PIT 抽查**：取 2 个历史样本，确认信号只用 ≤报价日 数据（无未来）。
- [ ] 波二若【待定】需调整，只改 params.WAVE2 + wave2.py，重跑即可。

---

## Self-Review（计划自检）

- **Spec 覆盖**：波二(Task3) ✓、抵抗(Task4) ✓、屏幕(Task6) ✓、特征(Task7) ✓、预筛(Task6 universe 参数) ✓、PIT(data_loader Task5 ≤date) ✓。
- **占位/待修点已显式标 ⚠**：data_loader 字段核对(Task5)、screen 的 stock_r 占位必修(Task6)、derive_strategy_signals pass 段(Task7)。这些都是"接真实 schema/数据"的实施动作，非设计占位。
- **类型一致**：wave2_signal(closes) / resist_score(stock_r, sector_r, market_r) 签名在 Task3/4 定义，Task6/7 调用一致。
- **波二【待定】**：参数集中 params.WAVE2，调整不波及 runner。

## 执行顺序建议

Task1-4（信号函数+单测，纯逻辑，可先全做完验证逻辑）→ Task5（接真实数据）→ Task6（屏幕，先跑通每日列表）→ Task7（特征接入）→ Task8（验证）。波二参数边跑边调。
