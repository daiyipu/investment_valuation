#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""特征业务解释词典 —— 评分卡审计/人工核查的基础。

评分卡上线前, 审核人员需看到每个指标的业务含义, 不能只看 AUC。
本模块提供:
  - EXPLICIT: 关键特征的精确业务解释(人工维护)。
  - explain(feature): 先查 EXPLICIT, 再按命名规则模式匹配(族前缀), 兜底"（待补）"。

新增特征时: 若命名规则已覆盖(如 turnover_*/sue_*/MACD_*)自动有解释;
  特殊语义再在 EXPLICIT 里补一条。
"""
import re

# ── 精确解释(人工维护; 优先级最高) ──
EXPLICIT = {
    # 11feat 生产评分卡
    'ROC_M_3': '月线3期收益率(月度动量); 高=近期月线走强',
    'MACD_HIST': '日线MACD柱(12,26,9的 DIF-DEA×2); 正且高=多头动能扩张',
    'ROC_W_3': '周线3期收益率(周度动量)',
    'k_KLEN': 'K线振幅 (日内最高-最低)/收盘; 大=波动剧烈',
    'MACD_W_HIST': '周线MACD柱; 中期多头动能',
    'BOLL_BW': '布林带宽度 (上轨-下轨)/均值; 大=波动率扩张',
    'ret_skew_60': '60日收益偏度; 正=偶有大涨, 负=偶有大跌',
    'vwap_dist': '收盘价相对VWAP偏离 (close/vwap-1); 正=收盘高于成交均价',
    '营收_3年斜率': '近3年营业收入增长斜率; 高=营收持续成长',
    'MACD_M_DEA': '月线MACD信号线(慢线); 趋势中期方向',
    '三浪_gain': '三浪策略一浪(启动浪)涨幅; 大=前段主升浪强',
    # 量价族
    'corr_ret_vol_5': '近5日 日收益-成交量 相关(量价协同/背离); 负=量价背离',
    'corr_ret_vol_20': '近20日 日收益-成交量 相关; 方正量价背离核心因子',
    'corr_ret_vol_60': '近60日 日收益-成交量 相关',
    'corr_close_vol_20': '近20日 收盘价-成交量 相关',
    'obv_slope_20': '20日OBV斜率(归一化净买入量); 高=资金累积',
    'pvt_slope_20': '20日PVT(量价趋势)斜率; 高=放量上涨累积(均值回归则反)',
    'mfi_14': '资金流向指标MFI(14)(量加权RSI, 0-100); 高=资金流入过热',
    'cmf_20': 'Chaikin资金流(20日 ΣCLV·量/总量, -1~1); 正=资金收集',
    'vol_surge_20': '当日量/近20日最大量; 高=放量(量在价先)',
    'vol_mom_5_20': '量能动量(5日均量/20日均量-1); 高=近期放量',
    'vol_ratio_20': '当日量/20日均量',
    'amount_ratio_20': '当日额/20日均额',
    'vwap_dist_old': '收盘/VWAP偏离',
    # 换手率族
    'turnover_mean_20': '近20日平均换手率; 高=交投活跃(投机)',
    'turnover_mean_60': '近60日平均换手率',
    'turnover_now': '最新换手率',
    'turnover_mom_20': '换手加速(最新/20日均-1)',
    'turnover_std_20': '近20日换手率标准差(交易稳定性, 华西); 高=换手剧烈波动',
    'turnover_skew_20': '近20日换手率偏度',
    'volratio_now': '最新量比(当日成交量/过去5日均量)',
    # SUE 业绩超预期族
    'sue_yoy': '报价日前最近一期已披露业绩的同比净利增速(PEAD代理)',
    'sue_zscore': '同比超预期的个股历史z-score(Latane-Jones SUE)',
    'sue_beat': '超自家指引=(实际/快报净利-预告中点)/|预告中点|; 正=超指引',
    'sue_recency_d': '最近披露距报价日天数(新鲜度); 近=PEAD强',
    'sue_yoy_mean3': '近3年报净利同比均值(盈利水平, 中长期稳定)',
    'sue_yoy_acc': '最新年报同比-上年同比(盈利加速度)',
    'sue_pos_streak': '连续盈利(YoY>0)年报期数(持续盈利能力)',
    'sue_up_trend': '近3年报同比改善趋势比例',
}

# ── 命名规则模式兜底(族前缀) ──
_PATTERNS = [
    (r'^MACD_[WM]?', 'MACD类动量指标'),
    (r'^RSI_', 'RSI相对强弱(超买超卖)'),
    (r'^KDJ_', 'KDJ随机指标'),
    (r'^BOLL_', '布林带指标'),
    (r'^ROC_', '收益率动量(roc)'),
    (r'^beta_', 'Beta系数(个股对市场/行业敏感度)'),
    (r'^smc_', '聪明钱SMC结构指标'),
    (r'^chip_', '筹码分布因子(cyq_chips)'),
    (r'^mf_', '资金流(moneyflow, 大单/主力净流入)'),
    (r'^nb_', '北向资金(hk_hold)持仓'),
    (r'^turnover_', '换手率族(daily_basic派生)'),
    (r'^sue_', '业绩超预期SUE族(forecast/express/income PIT)'),
    (r'^corr_ret_vol', '日收益-成交量相关(量价协同/背离)'),
    (r'^k_', 'K线形态因子'),
    (r'^ret_', '收益分布矩(偏度/峰度等)'),
    (r'FCF', '自由现金流衍生'),
    (r'NOPAT', '税后营业利润衍生'),
    (r'营收', '营业收入增长衍生'),
    (r'净利润', '净利润增长衍生'),
    (r'资本支出|capex', '资本支出衍生'),
    (r'行业', '行业估值/动量'),
    (r'同行', '同行业相对估值'),
    (r'区间收益|年化收益', '历史价格区间收益'),
    (r'MA\d', '均线位置/距离'),
]


def explain(feature):
    """特征→中文业务解释。先 EXPLICIT, 再模式匹配, 兜底待补。"""
    if feature in EXPLICIT:
        return EXPLICIT[feature]
    for pat, desc in _PATTERNS:
        if re.search(pat, str(feature)):
            return desc
    return '（待补业务解释）'


if __name__ == '__main__':
    import sys
    feats = sys.argv[1:] or list(EXPLICIT.keys())
    for f in feats:
        print(f'{f:<22} {explain(f)}')
