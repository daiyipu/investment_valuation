#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
特征剔除统一清单 (compare_selection.py / train_scorecard.py 共用)

集中管理"不纳入模型训练"的字段，避免各脚本各自维护、互不一致导致漂移。
分两类：
  1) LEAKAGE_PATTERNS : 泄漏/标识/时间戳/覆盖率 artifact —— 子串匹配
  2) BUSINESS_DROP    : 业务上无意义、样本期虚高相关的字段 —— 精确列名

新增剔除项只需在此文件追加，两个训练脚本自动同步生效。
"""


# 1) 泄漏 / 标识符 / 时间戳 / 数据覆盖率 artifact
#    子串匹配：列名包含任一模式即剔除
LEAKAGE_PATTERNS = [
    '报价日',       # 定增报价日及其衍生(报价日价格/报价日MA20/报价日_md 等) —— 时间泄漏
    '邀请日',       # 定增邀请日 —— 时间泄漏
    '最新交易日',   # 行情截止日 —— 时间泄漏
    '财报年份',     # 财报年份(日历年) —— 时间戳 artifact
    '数据天数',     # 行情数据覆盖天数 —— 采集 artifact，近常数
    'FCF年份',      # FCF年份_T/T1/T2/T3/T4 = 日历年(2015~2025)，不是现金流! —— 时间戳 artifact
    '_excel',       # Excel 来源标记列
    '_md',          # 元数据列
    # —— 多期限标签的原始收益列(目标变量本身): 7/1/3/6/12 个月涨跌幅/后价格, 含进特征即泄漏 ——
    '个月涨跌幅',
    '个月后价格',
    '周涨跌幅',      # 短线收益原始(return_1w/2w/4w → 1周/2周/4周涨跌幅; 目标变量)
    'excess_',       # 超额收益原始(excess_mkt/ind_*m; 目标变量, 非特征)
    '标签',          # 所有 标签_盈利_* / 标签_极性_灰度剔除_* 列(目标变量); 单一拦截点, 不再依赖各脚本 ad-hoc 过滤
]

# 2) 业务意义排除
#    精确列名匹配
BUSINESS_DROP = [
    '行业总天数',   # 行业指数历史天数，90% 样本挤在 666~675，数据采集 artifact，无经济学含义
    '成长能力_T-2', # 绝对水平值；delta 优先原则剔除。已验证 AUC 持平(逐步回归 0.728→0.727)，Lasso 会用 成长能力_delta_1y/营收_CAGR2 替代
    '净利润',       # 绝对金额(元)；绝对值跨公司/跨期不可比、损害泛化，不进评分卡。已验证剔除后 AUC 持平，相对指标(净利率/速动比率)替代
    # —— 个股绝对价格水平(元)：跨公司不可比、OOT 不泛化(Part A 抓到的)。保留相对版 price_vs_MA20/60/120/250 与 市场距离MA250 ——
    '当前价', 'MA20', 'MA30', 'MA60', 'MA120', 'MA250',
    '均价_all', '中位价', '价格标准差', '锁定价当前价', '报价日MA20',
    # —— 外部(询转)集整列为空 → median 填充成常数噪声，部署态无信号，剔除 ——
    'sw_l1_code', 'sw_l1_name', 'sw_l2_code', 'sw_l2_name', 'sw_l3_code', 'sw_l3_name',
    '行业PS', '定价方式',
    '评级_T-4', '评级_T-3', '评级_T-2', '评级_T-1', '评级_T',
    '综合趋势', '总分_趋势', '盈利能力_趋势', '成长能力_趋势',
    # —— 定增结构绝对原料(元/股/亿元, 跨公司不可比): 只用衍生比率(折价率/定增稀释率/募集市值比) ——
    '定增_发行价', '定增_增发数量', '定增_募资总额', '定增_发行前股本', '定增_发行后股本',
]


def get_excluded_columns(columns):
    """给定列名集合，返回应剔除的列名列表（两类合并，去重保序）。"""
    cols = list(columns)
    dropped = [c for c in cols if any(p in c for p in LEAKAGE_PATTERNS)]
    dropped += [c for c in cols if c in BUSINESS_DROP]
    seen = set()
    out = []
    for c in dropped:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out
