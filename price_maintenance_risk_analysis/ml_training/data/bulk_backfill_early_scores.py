#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bulk 回填 2010-2016 年度财务评分 —— 用 tushare fina_indicator_vip 按期批量接口。

架构(绕开 fina_indicator 三重瓶颈: 100行/次硬顶 + 500/min限频 + 逐成员调用):
  1. fina_indicator_vip(period=YYYY1231) 按期一次拉【全市场】财务指标(与 fina_indicator 同字段,
     vip 5000分支持按期批量; 替代原逐成员 ~45000 次)
  2. 按 (申万三级行业, 年) groupby 出【全部】行业年阈值(本地分位数) —— 指标源与既有
     2016-2026 数据(fina_indicator)一致, 无口径断层
  3. 本地评分(复用 FinancialScoringCard, 指标值直接喂入)
  4. 落库 company_annual_scores (ON DUPLICATE KEY UPDATE)

注: 不再用三张表算比率(避免 total_profit 等字段不对称坑); 指标接口直接给 roe/roa/margin 等。
阈值年份 = 各股报价日年(T)。
"""
import os, sys, time, argparse, threading
sys.path.insert(0, os.path.expanduser('~/github/EFAES'))
from src.web.config import DB_CONFIG, TUSHARE_TOKEN
from src.core.calculate_financial_score import FinancialScoringCard
import tushare as ts, pandas as pd, numpy as np
from concurrent.futures import ThreadPoolExecutor

PRO = ts.pro_api(TUSHARE_TOKEN)
NEG = {'debt_to_assets', 'int_to_talcap', 'debt_to_eqt'}
INDICATORS = ['inv_turn', 'ar_turn', 'ca_turn', 'assets_turn', 'netprofit_margin', 'grossprofit_margin',
              'roe', 'roe_dt', 'roa', 'npta', 'current_ratio', 'quick_ratio', 'cash_to_liqdebt',
              'cash_to_liqdebt_withinterest', 'debt_to_assets', 'int_to_talcap', 'debt_to_eqt', 'ebit_to_interest',
              'netprofit_yoy', 'dt_netprofit_yoy', 'roe_yoy', 'tr_yoy', 'or_yoy', 'equity_yoy', 'op_yoy', 'ebt_yoy',
              'rd_exp_ratio']
_VIP_SEM = threading.Semaphore(4)  # vip 接口并发上限


def _vip(fn, period):
    for attempt in range(3):
        try:
            with _VIP_SEM:
                d = getattr(PRO, fn)(period=period)
            if d is not None and not d.empty:
                return d
        except Exception as e:
            print(f'    {fn}({period}) 重试{attempt}: {str(e)[:80]}')
            time.sleep(1.0 * (attempt + 1))
    return None


def bulk_indicators(years):
    """fina_indicator_vip 按期拉全市场财务指标(与 fina_indicator 同字段, vip 支持按期批量)。
    过滤年报(end_date==period), 按 ts_code 去重(保留最新 ann_date)。返回含 report_year 的 DataFrame。"""
    frames = []
    for y in years:
        d = _vip('fina_indicator_vip', f'{y}1231')
        if d is None or d.empty:
            print(f'  fina_indicator_vip({y}): 空')
            continue
        d = d[d['end_date'].astype(str) == f'{y}1231'].copy()
        if d.empty:
            continue
        d = d.sort_values('ann_date', ascending=False).drop_duplicates('ts_code', keep='first')
        d['report_year'] = y
        frames.append(d)
        print(f'  fina_indicator_vip({y}): {len(d)} 公司')
    ratios = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if not ratios.empty:
        ratios['ts_code'] = ratios['ts_code'].astype(str)
        print(f'  合并: {len(ratios)} 行, {ratios["ts_code"].nunique()} 公司, '
              f'{ratios["report_year"].min()}~{ratios["report_year"].max()}年')
    return ratios


def industry_of(ts_code):
    """index_member_all 反查申万行业(支持并发, 退避重试)。返回名称+code(code 用于正向取成分)。"""
    for attempt in range(3):
        try:
            d = PRO.index_member_all(ts_code=ts_code)
            if d is not None and not d.empty:
                r = d.sort_values('in_date', ascending=False).iloc[0]
                if r.get('l3_name'):
                    return {'l3': r.get('l3_name'), 'l3_code': r.get('l3_code'),
                            'l2': r.get('l2_name'), 'l2_code': r.get('l2_code'),
                            'l1': r.get('l1_name'), 'l1_code': r.get('l1_code')}
                return None  # 非空但无三级 → 真无分类, 不重试
        except Exception:
            pass
        if attempt < 2:
            time.sleep(0.5 * (attempt + 1))
    return None


def members_by_code(code, level):
    """index_member_all 正向(按 l3_code/l2_code)取行业成分股 ts_code, 带重试。
    index_member_all 单接口双向: 反查用 ts_code, 正查用 l3_code/l2_code(注意用 code 不是 name,
    name 不被识别会返回全量)。并发友好, 替代 index_classify+index_member 两步。"""
    if not code:
        return []
    for attempt in range(3):
        try:
            d = PRO.index_member_all(**{f'{level}_code': code})
            if d is not None and not d.empty:
                return d['ts_code'].astype(str).unique().tolist()
            return []  # 非空但无成分 → 真无
        except Exception:
            if attempt < 2:
                time.sleep(0.6 * (attempt + 1))
    return []


def compute_thresholds(ratios, industry_members_map):
    """按 (行业, 年) 出阈值 {indicator: {极差/较差/中等/良好/优秀}}。负向指标反向。"""
    thresholds = {}
    for l3, members in industry_members_map.items():
        if not members:
            continue
        sub = ratios[ratios['ts_code'].isin(members)]
        for y, g in sub.groupby('report_year'):
            if len(g) < 5:
                continue
            t = {}
            for col in INDICATORS:
                if col not in g.columns:
                    continue
                vals = g[col].replace([np.inf, -np.inf], np.nan).dropna()
                if len(vals) < 5:
                    continue
                p10, p30, p50, p70, p90 = (vals.quantile(q) for q in (0.1, 0.3, 0.5, 0.7, 0.9))
                if pd.isna(p50):
                    continue
                if col in NEG:
                    t[col] = {'优秀': p10, '良好': p30, '中等': p50, '较差': p70, '极差': p90}
                else:
                    t[col] = {'极差': p10, '较差': p30, '中等': p50, '良好': p70, '优秀': p90}
            if len(t) >= 10:
                thresholds[(l3, int(y))] = t
    return thresholds


def score_stock(ts_code, t_year, ind_info, ratios, thr_l3, thr_l2, scorer):
    """本地评一只股 T-4..T 年。阈值三级→二级 回退(冷门三级样本不足时兜底)。"""
    results = {}
    l3 = ind_info.get('l3') if ind_info else None
    l2 = ind_info.get('l2') if ind_info else None
    for y in range(t_year - 4, t_year + 1):
        rrow = ratios[(ratios['ts_code'] == ts_code) & (ratios['report_year'] == y)]
        thr = thr_l3.get((l3, y)) or thr_l2.get((l2, y))
        if rrow.empty or not thr:
            continue
        ratios_dict = {k: v for k, v in rrow.iloc[0].to_dict().items()
                       if k not in ('ts_code', 'company_name', 'ann_date', 'end_date', 'report_year', 'used_average_values')}
        scorer.set_industry_thresholds(thr)
        scorer.load_company_data(ratios_dict, ts_code)
        sc = scorer.calculate_score()
        results[y] = {'总分': round(sc['总得分'], 2), '评级': scorer.get_rating(sc['总得分']),
                      '盈利能力': round(sc['维度得分'].get('盈利能力', 0), 2),
                      '成长能力': round(sc['维度得分'].get('成长能力', 0), 2),
                      '运营能力': round(sc['维度得分'].get('运营能力', 0), 2),
                      '偿债能力': round(sc['维度得分'].get('偿债能力', 0), 2)}
    return results


def save_to_db(stock_code, ind_info, results):
    import pymysql
    if not results:
        return
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')
    try:
        with conn.cursor() as cur:
            for year, r in results.items():
                cur.execute("""INSERT INTO company_annual_scores
                    (stock_code, report_year, total_score, rating, profitability, growth, operating, solvency,
                     industry_l1, industry_l2, industry_l3)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    ON DUPLICATE KEY UPDATE total_score=VALUES(total_score), rating=VALUES(rating),
                    profitability=VALUES(profitability), growth=VALUES(growth), operating=VALUES(operating),
                    solvency=VALUES(solvency), industry_l1=VALUES(industry_l1), industry_l2=VALUES(industry_l2),
                    industry_l3=VALUES(industry_l3)""",
                            (stock_code, year, r['总分'], r['评级'], r['盈利能力'], r['成长能力'], r['运营能力'], r['偿债能力'],
                             (ind_info or {}).get('l1'), (ind_info or {}).get('l2'), (ind_info or {}).get('l3')))
            conn.commit()
    finally:
        conn.close()


def main():
    ap = argparse.ArgumentParser(description='Bulk vip 回填 2010-2016 年度评分')
    ap.add_argument('excel', help='输入 Excel(股票代码/股票简称/报价日)')
    ap.add_argument('--limit', type=int, default=0, help='只处理前N只(0=全部)')
    args = ap.parse_args()

    df = pd.read_excel(args.excel)
    df = df.rename(columns={c: c.strip() for c in df.columns})
    placements = [(str(r['股票代码']).strip(), int(str(r['报价日']).strip()[:4]))
                  for _, r in df.iterrows() if pd.notna(r['股票代码'])]
    if args.limit:
        placements = placements[:args.limit]
    print(f'待回填: {len(placements)} 只股')

    # 年份范围(报价日 T-4..T 的并集; +1年做增长率基期)
    yrs = sorted({t - 4 for _, t in placements} | {t for _, t in placements})
    fetch_years = list(range(min(yrs) - 1, max(yrs) + 1))
    print(f'取数年份(含增长基期): {fetch_years[0]}~{fetch_years[-1]}')

    print('\n[1/3] fina_indicator_vip 按期拉全市场指标...')
    ratios = bulk_indicators(fetch_years)
    ratios['report_year'] = ratios['report_year'].astype(int)

    print('\n[2/3] 各股行业反查 + 三级/二级成分股 + 全部(行业,年)阈值...')
    unique_codes = sorted({c for c, _ in placements})
    with ThreadPoolExecutor(max_workers=8) as ex:
        ind_map = dict(ex.map(lambda c: (c, industry_of(c)), unique_codes))
    print(f'  {len(unique_codes)} 只唯一股 / {len(placements)} 条placement')
    unique_l3 = sorted({i['l3'] for i in ind_map.values() if i and i.get('l3')})
    unique_l2 = sorted({i['l2'] for i in ind_map.values() if i and i.get('l2')})
    print(f'  涉及 {len(unique_l3)} 个三级 / {len(unique_l2)} 个二级行业')

    def _code_for(name, name_key, code_key):
        return next((i.get(code_key) for i in ind_map.values() if i and i.get(name_key) == name), None)

    with ThreadPoolExecutor(max_workers=8) as ex:
        mm_l3 = dict(zip(unique_l3, ex.map(lambda n: members_by_code(_code_for(n, 'l3', 'l3_code'), 'l3'), unique_l3)))
        mm_l2 = dict(zip(unique_l2, ex.map(lambda n: members_by_code(_code_for(n, 'l2', 'l2_code'), 'l2'), unique_l2)))
    thr_l3 = compute_thresholds(ratios, mm_l3)
    thr_l2 = compute_thresholds(ratios, mm_l2)
    print(f'  阈值: 三级 {len(thr_l3)} / 二级 {len(thr_l2)} 个(行业,年)组合')

    print('\n[3/3] 本地评分 + 落库...')
    scorer = FinancialScoringCard(DB_CONFIG)
    ok = 0
    for i, (code, t_year) in enumerate(placements, 1):
        ind = ind_map.get(code)
        if not ind or not (ind.get('l3') or ind.get('l2')):
            continue
        results = score_stock(code, t_year, ind, ratios, thr_l3, thr_l2, scorer)
        if results:
            save_to_db(code, ind, results)
            ok += 1
        if i % 50 == 0 or i == len(placements):
            print(f'  [{i}/{len(placements)}] 已评分 {ok} 只')
    print(f'\n完成: {ok}/{len(placements)} 只股评分落库')


if __name__ == '__main__':
    main()
