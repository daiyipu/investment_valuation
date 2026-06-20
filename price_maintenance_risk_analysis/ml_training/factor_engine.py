#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""因子引擎 —— 声明式因子算子库, 从个股 OHLCV + 市场/行业收益(PIT 切片)计算因子。

参考 Microsoft Qlib Alpha158 思路(公式化因子), 补本项目缺失的两族:

  Beta 族(结构性、抗 regime 漂移 → 压 train-test gap):
    beta_mkt_{60,120,250}   个股对大盘的 β = Cov(Rs,Rm)/Var(Rm)
    beta_ind_120            个股对行业指数的 β
    idiovol_120             特质波动率 = std(Rs − β·Rm), 去市场成分后的个股风险

  Alpha158 缺失族(短中期价格行为信号):
    K线形态  KMID/KLEN/KMID2/KUP/KLOW/实体占比 (最新日, scale-free)
    技术指标 RSI(6/12/24) / KDJ(9,3,3) / MACD(12,26,9) / 布林%B+bandwidth
    量价     vol_ratio_20 / amount_ratio_20 / corr_close_vol_20 / vwap_dist
    滚动矩   skew/kurt(20/60) / roc(5/10/20/60)

所有因子纯函数; 入参为 numpy 数组(已 ≤报价日 切片, PIT 由调用方保证);
返回 {因子名: 值}。单因子异常填 NaN 不影响其余。
"""
import numpy as np
import pandas as pd

EPS = 1e-12


# ─────────────── 基础算子 ───────────────
def _ret(close):
    """对数收益(长度 N-1)。"""
    close = np.asarray(close, float)
    return np.diff(np.log(close))


def _sma(s, n):
    s = np.asarray(s, float)
    return float(np.nanmean(s[-n:])) if len(s) >= n else np.nan


def _std(s, n):
    s = np.asarray(s, float)
    return float(np.nanstd(s[-n:], ddof=0)) if len(s) >= n else np.nan


def _ema(s, span):
    s = pd.Series(np.asarray(s, float))
    return s.ewm(span=span, adjust=False).mean().values


def _rsi(close, n):
    close = np.asarray(close, float)
    if len(close) < n + 1:
        return np.nan
    diff = np.diff(close)
    gain = np.where(diff > 0, diff, 0.0)
    loss = np.where(diff < 0, -diff, 0.0)
    avg_gain = np.mean(gain[-n:]); avg_loss = np.mean(loss[-n:])
    if avg_loss < EPS:
        return 100.0
    rs = avg_gain / avg_loss
    return float(100 - 100 / (1 + rs))


def _macd(close):
    """返回 (dif, dea, hist)。"""
    close = np.asarray(close, float)
    if len(close) < 35:
        return np.nan, np.nan, np.nan
    ema12 = _ema(close, 12); ema26 = _ema(close, 26)
    dif = ema12 - ema26
    dea = _ema(dif, 9)
    hist = (dif - dea) * 2
    return float(dif[-1]), float(dea[-1]), float(hist[-1])


def _kdj(high, low, close, n=9, m1=3, m2=3):
    high = np.asarray(high, float); low = np.asarray(low, float); close = np.asarray(close, float)
    if len(close) < n:
        return np.nan, np.nan, np.nan
    hn = pd.Series(high).rolling(n, min_periods=1).max()
    ln = pd.Series(low).rolling(n, min_periods=1).min()
    rsv = (close - ln) / (hn - ln + EPS) * 100
    k = rsv.ewm(alpha=1.0 / m1, adjust=False).mean()
    d = k.ewm(alpha=1.0 / m2, adjust=False).mean()
    j = 3 * k - 2 * d
    return float(k.iloc[-1]), float(d.iloc[-1]), float(j.iloc[-1])


# ─────────────── 因子族 ───────────────
def kline_factors(o, h, l, c):
    """K线形态(最新日, scale-free 比率)。"""
    o, h, l, c = float(o[-1]), float(h[-1]), float(l[-1]), float(c[-1])
    rng = h - l
    body = c - o
    co_max, co_min = max(c, o), min(c, o)
    return {
        'k_KMID': body / (c + EPS),                 # (close-open)/close
        'k_KLEN': rng / (c + EPS),                  # (high-low)/close 振幅
        'k_KMID2': body / (rng + EPS),              # 实体/振幅
        'k_KUP': (h - co_max) / (rng + EPS),        # 上影/振幅
        'k_KLOW': (co_min - l) / (rng + EPS),       # 下影/振幅
        'k_BODY_RATIO': abs(body) / (rng + EPS),    # |实体|/振幅
    }


def tech_factors(c, h, l):
    """技术指标: RSI/KDJ/MACD/布林。"""
    out = {
        'RSI_6': _rsi(c, 6), 'RSI_12': _rsi(c, 12), 'RSI_24': _rsi(c, 24),
    }
    k, d, j = _kdj(h, l, c)
    out.update({'KDJ_K': k, 'KDJ_D': d, 'KDJ_J': j})
    dif, dea, hist = _macd(c)
    out.update({'MACD_DIF': dif, 'MACD_DEA': dea, 'MACD_HIST': hist})
    # 布林(20,2)
    if len(c) >= 20:
        ma = np.mean(c[-20:]); sd = np.std(c[-20:], ddof=0)
        up, lo = ma + 2 * sd, ma - 2 * sd
        out['BOLL_%B'] = float((c[-1] - lo) / (up - lo + EPS))   # 0=下轨 1=上轨
        out['BOLL_BW'] = float((up - lo) / (ma + EPS))           # 带宽
    else:
        out['BOLL_%B'] = out['BOLL_BW'] = np.nan
    return out


def volume_factors(c, h, l, v, amount):
    """量价族: 量比/额比/量价相关/VWAP偏离 + 量能动量(OBV/MFI/CMF/PVT)/多窗口收益-量相关。

    corr_ret_vol_n 为方正"量价协同/背离"核心(<0=价涨量缩的量价背离);
    OBV/CMF/PVT 捕捉资金累积方向; MFI 为量加权 RSI; vol_surge/mom 为"量在价先"。
    """
    c = np.asarray(c, float); h = np.asarray(h, float); l = np.asarray(l, float)
    v = np.asarray(v, float); amount = np.asarray(amount, float)
    out = {
        'vol_ratio_20': float(v[-1] / (_sma(v, 20) + EPS)) if len(v) >= 20 else np.nan,
        'amount_ratio_20': float(amount[-1] / (_sma(amount, 20) + EPS)) if len(amount) >= 20 else np.nan,
    }
    if len(c) >= 20:
        cc = pd.Series(c[-20:]); vv = pd.Series(v[-20:])
        out['corr_close_vol_20'] = float(cc.corr(vv)) if vv.std() > EPS else np.nan
    else:
        out['corr_close_vol_20'] = np.nan
    # VWAP = amount/vol; 偏离 = close/vwap - 1
    if len(v) and v[-1] > EPS and len(amount):
        out['vwap_dist'] = float(c[-1] / (amount[-1] / v[-1]) - 1)
    else:
        out['vwap_dist'] = np.nan

    # 多窗口 收益-成交量 相关(方正量价协同/背离核心: <0=量价背离)
    ret = _ret(c)
    rv = v[1:]
    for n in (5, 20, 60):
        if len(ret) >= n and len(rv) >= n:
            r_s = pd.Series(ret[-n:]); v_s = pd.Series(rv[-n:])
            out[f'corr_ret_vol_{n}'] = float(r_s.corr(v_s)) if v_s.std() > EPS else np.nan
        else:
            out[f'corr_ret_vol_{n}'] = np.nan

    # OBV 斜率(归一化净买入): 20 日 signed volume / 总量
    if len(c) >= 22:
        sign = np.sign(np.diff(c))
        obv = np.cumsum(sign * v[1:])
        out['obv_slope_20'] = float((obv[-1] - obv[-21]) / (np.sum(v[-20:]) + EPS))
    else:
        out['obv_slope_20'] = np.nan
    # PVT 斜率(归一化): cumsum(ret*vol) 变动 /(均价×总量)
    if len(c) >= 22:
        pvt = np.cumsum(ret * v[1:])
        denom = np.mean(c[-20:]) * np.sum(v[-20:]) + EPS
        out['pvt_slope_20'] = float((pvt[-1] - pvt[-21]) / denom)
    else:
        out['pvt_slope_20'] = np.nan
    # MFI(14): 量加权 RSI, typical_price=(h+l+c)/3
    if len(c) >= 15:
        tp = (h + l + c) / 3.0
        mf = tp * v
        dtp = np.diff(tp)
        pos = np.where(dtp > 0, mf[1:], 0.0); neg = np.where(dtp < 0, mf[1:], 0.0)
        pmf = np.sum(pos[-14:]); nmf = np.sum(neg[-14:])
        out['mfi_14'] = float(100 - 100 / (1 + pmf / (nmf + EPS))) if nmf > EPS else 100.0
    else:
        out['mfi_14'] = np.nan
    # CMF(20): Chaikin Money Flow = Σ(CLV·vol)/Σvol, CLV∈[-1,1]
    if len(c) >= 20:
        clv = ((c - l) - (h - c)) / (h - l + EPS)
        out['cmf_20'] = float(np.sum((clv * v)[-20:]) / (np.sum(v[-20:]) + EPS))
    else:
        out['cmf_20'] = np.nan
    # 量在价先: 当日量/近20日最大量; 量能动量=5日均量/20日均量-1
    if len(v) >= 20:
        out['vol_surge_20'] = float(v[-1] / (np.max(v[-20:]) + EPS))
        s5, s20 = _sma(v, 5), _sma(v, 20)
        out['vol_mom_5_20'] = float(s5 / (s20 + EPS) - 1)
    else:
        out['vol_surge_20'] = out['vol_mom_5_20'] = np.nan
    return out


def turnover_factors(turnover, vol_ratio):
    """换手率/量比族(daily_basic 派生, 报价日快照)。

    窗口扩到 120/250(≈半年/1年)匹配 7m 长期限(短窗 20/60 信号偏 1m/3m, 7m 衰减)。
    turnover_std/skew = 华西"交易稳定性"; turnover_trend_250_60 = 长期/中期换手趋势。
    """
    out = {}
    t = np.asarray(turnover, float) if (turnover is not None and len(turnover)) else None
    if t is not None:
        out['turnover_now'] = float(t[-1])
        s60 = None
        for w in (20, 60, 120, 250):
            sw = _sma(t, w)
            if w == 60:
                s60 = sw
            out[f'turnover_mean_{w}'] = sw
            if len(t) >= w:
                out[f'turnover_std_{w}'] = _std(t, w)
                s = pd.Series(t[-w:])
                out[f'turnover_skew_{w}'] = float(s.skew()) if s.std() > EPS else np.nan
            else:
                out[f'turnover_std_{w}'] = out[f'turnover_skew_{w}'] = np.nan
        out['turnover_mom_20'] = float(t[-1] / (_sma(t, 20) + EPS) - 1)
        out['turnover_trend_250_60'] = (float(_sma(t, 250) / (s60 + EPS) - 1)
                                        if len(t) >= 250 and s60 is not None else np.nan)
    else:
        for w in (20, 60, 120, 250):
            out[f'turnover_mean_{w}'] = out[f'turnover_std_{w}'] = out[f'turnover_skew_{w}'] = np.nan
        out['turnover_now'] = out['turnover_mom_20'] = out['turnover_trend_250_60'] = np.nan
    out['volratio_now'] = float(vol_ratio[-1]) if (vol_ratio is not None and len(vol_ratio)) else np.nan
    return out


def turnover_multiperiod(turnover_dates, turnover):
    """换手率周线(_W)/月线(_M)版: 日换手率按 W/M 重采样(期内日均)取近 N 根 mean/std。

    月线换手率配月级标签(7m), 尺度匹配(类比 SMC 多周期 +2.4pp: 周/月线 bar 配月标签)。
    turnover_dates 与 turnover 等长('YYYYMMDD'); 无则全 NaN。
    """
    out = {f'turnover_mean{t}': np.nan for t in ('_W', '_M')}
    out.update({f'turnover_std{t}': np.nan for t in ('_W', '_M')})
    if turnover_dates is None or turnover is None or len(turnover) == 0:
        return out
    try:
        idx = pd.to_datetime(pd.Series(turnover_dates).astype(str), format='%Y%m%d')
        s = pd.Series(np.asarray(turnover, float), index=idx).dropna()
        for rule, tag, nbars in (('W', '_W', 13), ('M', '_M', 6)):   # 近13周/近6月
            r = s.resample(rule).mean().dropna()                      # 每周/月内的日均换手率
            if len(r) >= 3:
                rec = r.tail(nbars)
                out[f'turnover_mean{tag}'] = float(rec.mean())
                out[f'turnover_std{tag}'] = float(rec.std())
    except Exception:
        pass
    return out


def moment_factors(c):
    """滚动矩: skew/kurt + 多窗口 ROC。"""
    r = pd.Series(_ret(c))
    out = {}
    for n in (20, 60):
        out[f'ret_skew_{n}'] = float(r[-n:].skew()) if len(r) >= n else np.nan
        out[f'ret_kurt_{n}'] = float(r[-n:].kurt()) if len(r) >= n else np.nan
    for n in (5, 10, 20, 60):
        out[f'ROC_{n}'] = float(c[-1] / c[-1 - n] - 1) if len(c) > n and c[-1 - n] > EPS else np.nan
    return out


def beta_factors(sret, mret, iret):
    """Beta 族: 个股-大盘/行业 β + 特质波动率。sret/mret/iret 已对齐(同长度同日期)。"""
    out = {}
    for n in (60, 120, 250):
        if len(sret) >= n and len(mret) >= n:
            s, m = sret[-n:], mret[-n:]
            vm = np.var(m, ddof=1)
            out[f'beta_mkt_{n}'] = float(np.cov(s, m, ddof=1)[0, 1] / vm) if vm > EPS else np.nan
        else:
            out[f'beta_mkt_{n}'] = np.nan
    # 行业 β (120)
    if len(sret) >= 120 and len(iret) >= 120:
        s, ind = sret[-120:], iret[-120:]
        vi = np.var(ind, ddof=1)
        out['beta_ind_120'] = float(np.cov(s, ind, ddof=1)[0, 1] / vi) if vi > EPS else np.nan
    else:
        out['beta_ind_120'] = np.nan
    # 特质波动率(120): 残差 = s − β_mkt_120·m 的 std
    bm = out.get('beta_mkt_120')
    if bm is not None and not np.isnan(bm) and len(sret) >= 120:
        resid = sret[-120:] - bm * mret[-120:]
        out['idiovol_120'] = float(np.std(resid, ddof=1))
    else:
        out['idiovol_120'] = np.nan
    return out


def smc_factors(o, h, l, c, N=60, M=20):
    """Smart Money Concepts 因子(OHLC, 报价日=末根K线快照)。纯机械代理变量。

    N/M 为 lookback 窗口(K线数); 日线默认 60/20, 周线 48/12, 月线 24/6(由调用方传)。
    注: SMC 本质主观/盘中, 此处是机械代理, 捕捉机构足迹的结构性快照:
      smc_premium_discount  末根收盘在近N根区间的位置(0折价~1溢价; 均值回归)
      smc_fvg_net           近M根看涨FVG−看跌FVG 计数(价格不平衡方向)
      smc_bos               末根收盘破近N-1根 swing高/低(+1/-1/0, 结构突破)
      smc_liq_sweep         末根影线破前swing但收盘回归(±1猎杀止损, 0否)
      smc_displacement      近M根最大实体/真实波幅(机构强动能)
      smc_ob_retest         末根是否回到近M根最强阳线前的反向K线区域(订单块再入场, 1/0)
    """
    keys = ('smc_premium_discount', 'smc_fvg_net', 'smc_bos', 'smc_liq_sweep',
            'smc_displacement', 'smc_ob_retest', 'smc_ote', 'smc_liqvoid')
    out = {k: np.nan for k in keys}
    o = np.asarray(o, float); h = np.asarray(h, float)
    l = np.asarray(l, float); c = np.asarray(c, float)
    n = len(c)
    if n < max(25, M + 2):
        return out
    hN, lN = h[max(0, n-N):], l[max(0, n-N):]
    hM, lM, cM, oM = h[max(0, n-M):], l[max(0, n-M):], c[max(0, n-M):], o[max(0, n-M):]
    rng = hN.max() - lN.min()
    if rng <= 0:
        return out

    out['smc_premium_discount'] = (c[-1] - lN.min()) / rng

    bull = bear = 0
    for t in range(len(hM) - 2):
        if lM[t + 2] > hM[t]:
            bull += 1
        if hM[t + 2] < lM[t]:
            bear += 1
    out['smc_fvg_net'] = float(bull - bear)

    prev_hi = hN[:-1].max() if len(hN) > 1 else hN.max()
    prev_lo = lN[:-1].min() if len(lN) > 1 else lN.min()
    if c[-1] > prev_hi:
        out['smc_bos'] = 1.0
    elif c[-1] < prev_lo:
        out['smc_bos'] = -1.0
    else:
        out['smc_bos'] = 0.0

    if h[-1] > prev_hi and c[-1] < prev_hi:
        out['smc_liq_sweep'] = -1.0
    elif l[-1] < prev_lo and c[-1] > prev_lo:
        out['smc_liq_sweep'] = 1.0
    else:
        out['smc_liq_sweep'] = 0.0

    body = np.abs(cM - oM)
    tr = np.maximum.reduce([hM - lM, np.abs(hM - oM), np.abs(lM - oM)])
    atr = float(tr.mean())
    out['smc_displacement'] = float(body.max() / atr) if atr > 0 else np.nan

    if len(oM) >= 3:
        up_body = np.where(cM > oM, body, 0)
        bi = int(np.argmax(up_body))
        if bi > 0 and oM[bi - 1] > cM[bi - 1]:    # 强阳线前是阴线 = 经典 OB
            seg = [lM[bi - 1], oM[bi - 1], cM[bi - 1]]
            ob_lo, ob_hi = min(seg), max(max(seg), hM[bi - 1])
            out['smc_ob_retest'] = 1.0 if ob_lo <= c[-1] <= ob_hi else 0.0
        else:
            out['smc_ob_retest'] = 0.0

    # smc_ote: OTE 最优入场区(fib 0.62-0.79 回撤带; 价格落入=1, 均值回归入场)
    ote_lo = lN.min() + 0.62 * rng
    ote_hi = lN.min() + 0.79 * rng
    out['smc_ote'] = 1.0 if ote_lo <= c[-1] <= ote_hi else 0.0
    # smc_liqvoid: 流动性真空(近M日大实体 |c-o|>1.5*ATR 视为跳空真空; 上下方向计数差)
    if atr > 0:
        vmask = body > 1.5 * atr
        out['smc_liqvoid'] = float(np.where(cM[vmask] > oM[vmask], 1, -1).sum()) if vmask.any() else 0.0
    else:
        out['smc_liqvoid'] = 0.0
    return out


def smc_factors_multiperiod(dates, o, h, l, c):
    """SMC 因子的周线(_W)/月线(_M)版, 匹配中期标签(1m/3m/7m)。
    日线 SMC 配短标签(1w/2w); 周/月线 SMC 配月标签。lookback: W=48/12, M=24/6。"""
    out = {}
    keys = ('smc_premium_discount', 'smc_fvg_net', 'smc_bos', 'smc_liq_sweep',
            'smc_displacement', 'smc_ob_retest', 'smc_ote', 'smc_liqvoid')
    for rule, tag, N, M in (('W', '_W', 48, 12), ('M', '_M', 24, 6)):
        try:
            ow, hw, lw, cw = _resample_ohlc(dates, o, h, l, c, rule)
            if len(cw) < max(25, M + 2):
                for k in keys:
                    out[f'{k}{tag}'] = np.nan
                continue
            f = smc_factors(ow, hw, lw, cw, N=N, M=M)
            for k in keys:
                out[f'{k}{tag}'] = f.get(k, np.nan)
        except Exception:
            for k in keys:
                out[f'{k}{tag}'] = np.nan
    return out


def compute_factors(o, h, l, c, v, amount, sret, mret, iret, dates=None,
                    turnover=None, vol_ratio=None, turnover_dates=None):
    """主入口: 给定 OHLCV(同日期升序) + 对齐的三序列收益, 返回全部因子 dict。
    dates 给定时额外计算周/月线多周期因子(慢趋势)。
    turnover/vol_ratio(daily_basic 派生, ≤报价日 升序)给定时算换手率族;
    turnover_dates 给定时额外算周/月线换手率(配月级标签)。
    单族异常不影响其余(每族 try/except)。"""
    out = {}
    families = [
        (kline_factors, (o, h, l, c)),
        (tech_factors, (c, h, l)),
        (volume_factors, (c, h, l, v, amount)),
        (moment_factors, (c,)),
        (beta_factors, (sret, mret, iret)),
        (smc_factors, (o, h, l, c)),
        (turnover_factors, (turnover, vol_ratio)),
        (turnover_multiperiod, (turnover_dates, turnover)),
    ]
    if dates is not None:
        families.append((multiperiod_factors, (dates, o, h, l, c)))
        families.append((smc_factors_multiperiod, (dates, o, h, l, c)))
    for fn, args in families:
        try:
            out.update(fn(*args))
        except Exception:
            for k in _FACTOR_NAMES_OF.get(fn.__name__, []):
                out.setdefault(k, np.nan)
    return out


def _resample_ohlc(dates, o, h, l, c, rule):
    """日线 OHLC 重采样为周('W')/月('M') OHLC(按日历周/月聚合)。返回 (o,h,l,c) np 数组。"""
    idx = pd.to_datetime(pd.Series(dates).astype(str), format='%Y%m%d')
    dfx = pd.DataFrame({'o': o, 'h': h, 'l': l, 'c': c}, index=idx).dropna()
    r = dfx.resample(rule).agg({'o': 'first', 'h': 'max', 'l': 'min', 'c': 'last'}).dropna()
    return r['o'].values, r['h'].values, r['l'].values, r['c'].values


def multiperiod_factors(dates, o, h, l, c):
    """周/月线多周期因子(慢趋势, 对中长期更贴合)。日线重采样到 W/M 后算 RSI/MACD/KDJ/BOLL/ROC。
    月线 MACD 需 ~35 个月 → 调用方应传 ≥800 日线切片。"""
    out = {}
    for rule, tag in [('W', 'W'), ('M', 'M')]:
        try:
            ro, rh, rl, rc = _resample_ohlc(dates, o, h, l, c, rule)
            if len(rc) < 10:
                for k in _FACTOR_NAMES_OF['multiperiod_factors'] if tag == 'W' else []:
                    if k.endswith('_W'):
                        out[k] = np.nan
                continue
            out[f'RSI_{tag}'] = _rsi(rc, 6) if len(rc) >= 7 else np.nan
            if len(rc) >= 35:
                dif, dea, hist = _macd(rc)
                out[f'MACD_{tag}_DIF'] = dif; out[f'MACD_{tag}_DEA'] = dea; out[f'MACD_{tag}_HIST'] = hist
            else:
                out[f'MACD_{tag}_DIF'] = out[f'MACD_{tag}_DEA'] = out[f'MACD_{tag}_HIST'] = np.nan
            k, d, j = _kdj(rh, rl, rc)
            out[f'KDJ_{tag}_K'] = k; out[f'KDJ_{tag}_D'] = d; out[f'KDJ_{tag}_J'] = j
            if len(rc) >= 20:
                ma = np.mean(rc[-20:]); sd = np.std(rc[-20:], ddof=0)
                up, lo = ma + 2 * sd, ma - 2 * sd
                out[f'BOLL_{tag}_pctB'] = float((rc[-1] - lo) / (up - lo + EPS))
                out[f'BOLL_{tag}_BW'] = float((up - lo) / (ma + EPS))
            else:
                out[f'BOLL_{tag}_pctB'] = out[f'BOLL_{tag}_BW'] = np.nan
            out[f'ROC_{tag}_1'] = float(rc[-1] / rc[-2] - 1) if len(rc) > 2 and rc[-2] > EPS else np.nan
            out[f'ROC_{tag}_3'] = float(rc[-1] / rc[-4] - 1) if len(rc) > 4 and rc[-4] > EPS else np.nan
        except Exception:
            for k in _FACTOR_NAMES_OF['multiperiod_factors']:
                if k.endswith(f'_{tag}'):
                    out.setdefault(k, np.nan)
    return out


_FACTOR_NAMES_OF = {
    'kline_factors': ['k_KMID', 'k_KLEN', 'k_KMID2', 'k_KUP', 'k_KLOW', 'k_BODY_RATIO'],
    'tech_factors': ['RSI_6', 'RSI_12', 'RSI_24', 'KDJ_K', 'KDJ_D', 'KDJ_J',
                     'MACD_DIF', 'MACD_DEA', 'MACD_HIST', 'BOLL_%B', 'BOLL_BW'],
    'volume_factors': ['vol_ratio_20', 'amount_ratio_20', 'corr_close_vol_20', 'vwap_dist',
                       'corr_ret_vol_5', 'corr_ret_vol_20', 'corr_ret_vol_60',
                       'obv_slope_20', 'pvt_slope_20', 'mfi_14', 'cmf_20',
                       'vol_surge_20', 'vol_mom_5_20'],
    'turnover_factors': ['turnover_now', 'turnover_mom_20', 'turnover_trend_250_60', 'volratio_now']
                       + [f'{s}_{w}' for s in ('turnover_mean', 'turnover_std', 'turnover_skew')
                          for w in (20, 60, 120, 250)],
    'turnover_multiperiod': ['turnover_mean_W', 'turnover_std_W', 'turnover_mean_M', 'turnover_std_M'],
    'moment_factors': ['ret_skew_20', 'ret_kurt_20', 'ret_skew_60', 'ret_kurt_60',
                       'ROC_5', 'ROC_10', 'ROC_20', 'ROC_60'],
    'beta_factors': ['beta_mkt_60', 'beta_mkt_120', 'beta_mkt_250', 'beta_ind_120', 'idiovol_120'],
    'smc_factors': ['smc_premium_discount', 'smc_fvg_net', 'smc_bos', 'smc_liq_sweep',
                    'smc_displacement', 'smc_ob_retest', 'smc_ote', 'smc_liqvoid'],
    'smc_factors_multiperiod': [f'{k}{t}' for t in ('_W', '_M') for k in
                                ('smc_premium_discount', 'smc_fvg_net', 'smc_bos',
                                 'smc_liq_sweep', 'smc_displacement', 'smc_ob_retest',
                                 'smc_ote', 'smc_liqvoid')],
    'multiperiod_factors': [f'{x}_{t}' for t in ('W', 'M') for x in
                            ('RSI', 'MACD_DIF', 'MACD_DEA', 'MACD_HIST', 'KDJ_K', 'KDJ_D', 'KDJ_J',
                             'BOLL_pctB', 'BOLL_BW', 'ROC_1', 'ROC_3')],
}

ALL_FACTOR_NAMES = [n for v in _FACTOR_NAMES_OF.values() for n in v]


if __name__ == '__main__':
    # 冒烟: 合成一段 OHLCV + 收益
    np.random.seed(0)
    n = 300
    c = 10 * np.exp(np.cumsum(np.random.randn(n) * 0.02))
    o = c * (1 + np.random.randn(n) * 0.005)
    h = np.maximum(o, c) * (1 + np.abs(np.random.randn(n)) * 0.005)
    l = np.minimum(o, c) * (1 - np.abs(np.random.randn(n)) * 0.005)
    v = np.abs(np.random.randn(n)) * 1e6 + 1e5
    amt = c * v
    turnover = np.abs(np.random.randn(n)) * 0.5 + 1.0
    vol_ratio = np.abs(np.random.randn(n)) * 0.5 + 1.0
    sret = _ret(c); mret = _ret(c * 0.8 + np.random.randn(n) * 0.01); iret = _ret(c * 0.9)
    f = compute_factors(o, h, l, c, v, amt, sret, mret, iret, turnover=turnover, vol_ratio=vol_ratio)
    print(f'因子数: {len(f)} (期望 {len(ALL_FACTOR_NAMES)})')
    for k in sorted(f):
        print(f'  {k:<22} {f[k]:.4f}' if isinstance(f[k], float) and np.isfinite(f[k]) else f'  {k:<22} {f[k]}')
