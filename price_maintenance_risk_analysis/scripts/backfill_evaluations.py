#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
历史定增评估数据回填脚本（Step 6）

把带标签的历史 scored Excel 灌入两张 MySQL 表：
  - placement_evaluation    （每行一次定增：报价日/决策/子场景/趋势/7个月涨跌幅标签…）
  - company_annual_scores   （每只股票×每年：总分/评级/盈利能力/成长能力）

数据源: 带标签的 scored Excel（含「报价日」「7个月后涨跌幅」「总分_T-4..T」等列）。
两表均为 upsert（UNIQUE(stock_code,issue_date) / UNIQUE(stock_code,report_year)），可重复运行。

用法:
    python scripts/backfill_evaluations.py --labeled \\
        "/path/to/batch_screening_result_260609_意米定增_scored.xlsx" \\
        "/path/to/batch_screening_result_260610_意米定增_scored.xlsx"

格式约定（已与 DB 既有表对齐，无需手动转换）:
    股票代码   : '300164.SZ'            原样（DB market_data 等也用 .SZ/.SH 后缀）
    报价日     : 数字 20200109.0        → '20200109' 字符串（匹配 issue_date_locked.issue_date）
    溢价率     : '-20.00%'              → -0.2（小数；export_features 的 Excel 加载器也是 /100）
    7个月后涨跌: '-24.91%'              → -24.91（百分比数值；标签用 >0/>-10/>-20 比较）
"""

import os
import sys
import argparse
import pandas as pd

# 把项目根加入 sys.path，以便 from utils.db_manager import ValuationDB
_PROJ_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJ_ROOT not in sys.path:
    sys.path.insert(0, _PROJ_ROOT)

from utils.db_manager import ValuationDB

# 子场景列: Excel列名 → placement_evaluation 列名
SUB_SCENARIO_MAP = {
    '市场指数': 'sub_market_index',
    '行业PE': 'sub_industry_pe',
    '个股PE': 'sub_stock_pe',
    'DCF估值': 'sub_dcf',
    '修正PE估值': 'sub_adj_pe',
    '参数构造': 'sub_param_build',
    '蒙特卡洛': 'sub_monte_carlo',
    '反向推算': 'sub_reverse_calc',
}

RELATIVE_LABELS = ['T-4', 'T-3', 'T-2', 'T-1', 'T']


def _to_float(v):
    """安全转 float；空/'-' → None。不剥离 %。"""
    if v is None:
        return None
    s = str(v).strip()
    if s == '' or s == '-' or s.lower() == 'nan':
        return None
    try:
        return float(s.replace(',', ''))
    except (ValueError, TypeError):
        return None


def _pct_to_decimal(v):
    """百分比字符串 → 小数: '-20.00%' → -0.2；'-' → None。"""
    if v is None:
        return None
    s = str(v).strip().replace('%', '').replace('+', '').replace(' ', '')
    if s == '' or s == '-' or s.lower() == 'nan':
        return None
    try:
        return float(s) / 100.0
    except (ValueError, TypeError):
        return None


def _ret_to_pct(v):
    """涨跌幅百分比字符串 → 数值: '-24.91%' → -24.91（保留百分比数值，非小数）。"""
    if v is None:
        return None
    s = str(v).strip().replace('%', '').replace('+', '').replace(' ', '')
    if s == '' or s == '-' or s.lower() == 'nan':
        return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _issue_date_str(v):
    """报价日(数字 20200109.0 或 '20200109') → '20200109' 字符串。"""
    if v is None:
        return None
    try:
        return str(int(float(v)))
    except (ValueError, TypeError):
        s = str(v).strip().replace('-', '').replace('/', '')
        return s if s else None


def _sub_flag(v):
    """子场景值: 含'✓' → 1，否则 → 0。"""
    return 1 if (v is not None and '✓' in str(v)) else 0


def _row_to_placement_eval(row, batch_id):
    """单行 → placement_evaluation dict（供 save_placement_evaluation）。"""
    return {
        'stock_code': str(row.get('股票代码', '')).strip(),
        'stock_name': str(row.get('股票简称', '') or ''),
        'batch_id': batch_id,
        'issue_date': _issue_date_str(row.get('报价日')),
        'issue_date_price': _to_float(row.get('报价日价格')),
        'valid_thresholds': int(_to_float(row.get('有效阈值数')) or 0) if _to_float(row.get('有效阈值数')) is not None else None,
        'premium_min': _pct_to_decimal(row.get('溢价率下限')),
        'premium_max': _pct_to_decimal(row.get('溢价率上限')),
        'decision': str(row.get('定增决策', '') or ''),
        # 子场景
        **{db_col: _sub_flag(row.get(excel_col)) for excel_col, db_col in SUB_SCENARIO_MAP.items()},
        # 行业
        'industry_l1': str(row.get('一级行业', '') or ''),
        'industry_l2': str(row.get('二级行业', '') or ''),
        'industry_l3': str(row.get('三级行业', '') or ''),
        # 趋势
        'total_slope': _to_float(row.get('总分_斜率')),
        'total_trend': str(row.get('总分_趋势', '') or ''),
        'profit_slope': _to_float(row.get('盈利能力_斜率')),
        'profit_trend': str(row.get('盈利能力_趋势', '') or ''),
        'growth_slope': _to_float(row.get('成长能力_斜率')),
        'growth_trend': str(row.get('成长能力_趋势', '') or ''),
        'combined_trend': str(row.get('综合趋势', '') or ''),
        # 标签
        'return_7m': _ret_to_pct(row.get('7个月后涨跌幅')),
        'price_7m': _to_float(row.get('7个月后价格')),
        'final_conclusion': str(row.get('最终结论', '') or ''),
    }


def _row_to_annual_scores(row):
    """单行 → {report_year: score_dict}（供 save_annual_scores）。

    评分列用相对年份 T-4..T，按「报价日」年份回溯成绝对年份:
        T-4 → base_year-4, ..., T → base_year
    （与 export_features.load_scored_features_from_db 的反查逻辑一致）
    """
    issue_date = _issue_date_str(row.get('报价日'))
    if not issue_date or len(issue_date) < 4:
        return {}
    try:
        base_year = int(issue_date[:4])
    except ValueError:
        return {}

    scores_by_year = {}
    industry = {
        'industry_l1': str(row.get('一级行业', '') or ''),
        'industry_l2': str(row.get('二级行业', '') or ''),
        'industry_l3': str(row.get('三级行业', '') or ''),
    }
    stock_name = str(row.get('股票简称', '') or '')

    for i, label in enumerate(RELATIVE_LABELS):
        year = base_year - 4 + i
        total = _to_float(row.get(f'总分_{label}'))
        # 该行该年若无评分则跳过
        if total is None:
            continue
        scores_by_year[year] = {
            'stock_name': stock_name,
            'total_score': total,
            'rating': str(row.get(f'评级_{label}', '') or ''),
            'profitability': _to_float(row.get(f'盈利能力_{label}')),
            'growth': _to_float(row.get(f'成长能力_{label}')),
            'operating': None,   # Excel 无此维度
            'solvency': None,    # Excel 无此维度
            **industry,
        }
    return scores_by_year


def backfill_one_file(db, path, batch_id):
    """回填单个 Excel 文件，返回 (n_eval, n_scores, n_skipped)。"""
    if not os.path.exists(path):
        print(f'  ✗ 文件不存在: {path}')
        return 0, 0, 0

    df = pd.read_excel(path, sheet_name=0)
    print(f'\n  📄 {os.path.basename(path)}: {len(df)} 行 × {len(df.columns)} 列')
    print(f'     报价日 非空: {df["报价日"].notna().sum() if "报价日" in df.columns else 0}'
          f' | 7个月后涨跌幅 非空: {df["7个月后涨跌幅"].notna().sum() if "7个月后涨跌幅" in df.columns else 0}')

    n_eval = n_scores = n_skipped = 0
    for _, row in df.iterrows():
        code = str(row.get('股票代码', '')).strip()
        if not code:
            n_skipped += 1
            continue

        # 1. placement_evaluation
        pe = _row_to_placement_eval(row, batch_id)
        if pe['issue_date']:
            try:
                db.save_placement_evaluation(pe)
                n_eval += 1
            except Exception as e:
                print(f'     ⚠ {code} placement 落库失败: {e}')

        # 2. company_annual_scores
        scores_by_year = _row_to_annual_scores(row)
        if scores_by_year:
            try:
                db.save_annual_scores(code, scores_by_year, batch_id=batch_id)
                n_scores += len(scores_by_year)
            except Exception as e:
                print(f'     ⚠ {code} annual 落库失败: {e}')

    print(f'     ✅ placement_evaluation: {n_eval} 条 | company_annual_scores: {n_scores} 条 | 跳过: {n_skipped}')
    return n_eval, n_scores, n_skipped


def main():
    parser = argparse.ArgumentParser(description='历史定增评估数据回填 → placement_evaluation + company_annual_scores')
    parser.add_argument('--labeled', nargs='+', required=True,
                        help='带标签的 scored Excel 路径（可多个）')
    parser.add_argument('--batch-id', default=None,
                        help='批次ID标记（默认 backfill_YYYYMMDD）')
    args = parser.parse_args()

    # 批次ID：用文件修改时间近似日期（避免在脚本里取系统时间的不确定性）
    batch_id = args.batch_id or 'backfill_hist'

    db = ValuationDB()
    print(f'回填目标: investment_valuation.{ "(placement_evaluation, company_annual_scores)" }')
    print(f'批次ID: {batch_id}')
    print(f'待回填文件: {len(args.labeled)} 个')

    total_eval = total_scores = 0
    for path in args.labeled:
        path = os.path.expanduser(path)
        n_eval, n_scores, _ = backfill_one_file(db, path, batch_id)
        total_eval += n_eval
        total_scores += n_scores

    print('\n' + '=' * 60)
    print(f'🎉 回填完成: placement_evaluation 累计写入 {total_eval} 条, '
          f'company_annual_scores 累计写入 {total_scores} 条')
    print('   （均为 upsert，重复行已更新而非重复插入）')
    print('=' * 60)


if __name__ == '__main__':
    main()
