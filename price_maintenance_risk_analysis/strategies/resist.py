"""抵抗信号函数(纯函数, 无 IO)。

入参为**日收益率序列**(stock_r, sector_r, market_r, 等长, 对齐日期, 末项=当日)。
data_loader 负责取价→算收益对齐。
逻辑: A.相关性背离(个股-行业, 多窗5/10/20) + B.下跌波段抗跌方向 → 综合。
"""
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
