#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增特征导出脚本 - 从多个数据源汇总特征供ML训练

数据源:
  1. valuation.db (SQLite) - 行情特征/估值/FCF/筛选决策
  2. scored Excel - 财务评分/子场景/7个月涨跌幅
  3. EFAES MySQL - 财务比率(23指标，仅部分公司有)

用法:
    python ml_training/export_features.py <scored_excel_path> [--output features.parquet]

输出:
    ml_training/data/features.parquet - 一行一只股票，所有特征平铺
"""

import sys
import os
import sqlite3  # 兼容旧导入，实际用pymysql
import pymysql
import argparse
import numpy as np
import pandas as pd


def load_sqlite_features(db_path, stock_codes):
    """从MySQL加载行情/估值/FCF等特征"""
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4',
                           cursorclass=pymysql.cursors.DictCursor)
    cur = conn.cursor()

    features = []
    for code in stock_codes:
        f = {'股票代码': code}

        # 1. market_data 行情特征
        row = cur.execute('SELECT * FROM market_data WHERE stock_code=%s', (code,)).fetchone()
        if row:
            d = row
            # 波动率×4窗口
            for w in ['20d', '60d', '120d', '250d']:
                f[f'波动率_{w}'] = d.get(f'volatility_{w}')
                f[f'年化收益_{w}'] = d.get(f'annual_return_{w}')
                f[f'区间收益_{w}'] = d.get(f'period_return_{w}')
                f[f'胜率_{w}'] = d.get(f'win_rate_{w}')
            # 均线
            for ma in ['20', '30', '60', '120', '250']:
                f[f'MA{ma}'] = d.get(f'ma_{ma}')
            f['当前价'] = d.get('current_price')
            f['漂移率'] = d.get('drift')
            f['波动率'] = d.get('volatility')
            f['换手率'] = d.get('market_turnover')
            f['数据天数'] = d.get('total_days')
            f['schema_version'] = d.get('schema_version')

        # 2. relative_valuation 估值特征
        row = cur.execute('SELECT * FROM relative_valuation WHERE stock_code=%s', (code,)).fetchone()
        if row:
            d = row
            f['个股PE'] = d.get('current_pe')
            f['个股PB'] = d.get('current_pb')
            f['个股PS'] = d.get('current_ps')
            f['行业PE'] = d.get('sw_index_pe')
            f['行业PB'] = d.get('sw_index_pb')
            f['行业PS'] = d.get('sw_index_ps')
            f['行业代码'] = d.get('target_index_code', '')
            f['行业名称'] = d.get('target_industry_l3', '')

        # 3. historical_fcf FCF特征(取最近5年)
        rows = cur.execute(
            'SELECT * FROM historical_fcf WHERE stock_code=%s ORDER BY year DESC LIMIT 5', (code,)
        ).fetchall()
        for i, r in enumerate(rows):
            d = dict(r)
            suffix = f'_T{i}' if i > 0 else '_T'  # T=最近年, T1=前1年...
            f[f'营收{suffix}'] = d.get('revenue')
            f[f'NOPAT{suffix}'] = d.get('nopat')
            f[f'FCF{suffix}'] = d.get('fcf')
            f[f'FCF年份{suffix}'] = d.get('year')

        # 4. issue_date_locked 锁定价
        row = cur.execute('SELECT * FROM issue_date_locked WHERE stock_code=%s', (code,)).fetchone()
        if row:
            d = row
            f['报价日'] = d.get('issue_date')
            f['报价日价格'] = d.get('issue_date_price')
            f['报价日MA20'] = d.get('ma_20')

        # 5. placement_params 定增参数
        row = cur.execute('SELECT * FROM placement_params WHERE stock_code=%s', (code,)).fetchone()
        if row:
            d = row
            f['发行价'] = d.get('issue_price')
            f['锁定期'] = d.get('lockup_period')
            f['净债务'] = d.get('total_debt')
            f['净利润'] = d.get('net_income')
            f['净资产负债表'] = d.get('net_assets')

        # 6. screening_results 最新筛选结果
        row = cur.execute(
            'SELECT * FROM screening_results WHERE stock_code=%s ORDER BY id DESC LIMIT 1', (code,)
        ).fetchone()
        if row:
            d = row
            f['溢价率下限'] = d.get('premium_min')
            f['溢价率上限'] = d.get('premium_max')
            f['有效阈值数'] = d.get('valid_thresholds')
            f['step1通过'] = d.get('step1_pass')
            f['step2通过'] = d.get('step2_pass')
            f['step3通过'] = d.get('step3_pass')
            f['定增决策'] = d.get('decision')

        features.append(f)

    conn.close()
    return pd.DataFrame(features)


def load_scored_features(excel_path):
    """从scored Excel加载财务评分/子场景/涨跌幅"""
    df = pd.read_excel(excel_path, sheet_name='Sheet1')

    result = pd.DataFrame()
    result['股票代码'] = df['股票代码']
    result['股票简称'] = df.get('股票简称', '')

    # 财务评分(按相对年份列)
    score_cols = [c for c in df.columns if c.startswith('总分_') and c[3:].isdigit()]
    if score_cols:
        for prefix in ['总分', '评级', '盈利能力', '成长能力']:
            for c in df.columns:
                if c.startswith(f'{prefix}_') and c[len(prefix)+1:] in ['T-4','T-3','T-2','T-1','T']:
                    result[c] = df[c]
                elif c.startswith(f'{prefix}_') and c[len(prefix)+1:].isdigit():
                    result[c] = df[c]  # 兼容旧格式(绝对年份)

    # 趋势
    for c in ['总分_斜率', '总分_趋势', '盈利能力_斜率', '盈利能力_趋势',
              '成长能力_斜率', '成长能力_趋势', '综合趋势']:
        if c in df.columns:
            result[c] = df[c]

    # 子场景
    for c in ['市场指数', '行业PE', '个股PE', 'DCF估值', '修正PE估值',
              '参数构造', '蒙特卡洛', '反向推算']:
        if c in df.columns:
            result[f'{c}_通过'] = df[c].astype(str).str.contains('✓').astype(int)

    # 行业
    for c in ['一级行业', '二级行业', '三级行业']:
        if c in df.columns:
            result[c] = df[c]

    # 定增决策
    if '定增决策' in df.columns:
        result['定增建议参与'] = df['定增决策'].astype(str).str.contains('建议参与').astype(int)

    # 标签: 7个月涨跌幅
    if '7个月后涨跌幅' in df.columns:
        ret = pd.to_numeric(
            df['7个月后涨跌幅'].astype(str).str.replace('%', '').str.replace('+', ''),
            errors='coerce'
        )
        result['7个月涨跌幅'] = ret
        result['标签_盈利_0'] = (ret > 0).astype(int)          # 盈利(>0)
        result['标签_盈利_-10'] = (ret > -0.10).astype(int)     # 盈利(>-10%)
        result['标签_盈利_-20'] = (ret > -0.20).astype(int)     # 盈利(>-20%)

    # 最终结论
    if '最终结论' in df.columns:
        result['最终结论'] = df['最终结论']

    # 报价日
    if '报价日' in df.columns:
        result['报价日_excel'] = df['报价日']

    return result


def load_mysql_features(stock_codes):
    """从EFAES MySQL加载财务比率(23指标)"""
    try:
        import pymysql
        conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                               database='fund_risk_control')
        # 只取有的公司
        codes_str = ','.join([f"'{c}'" for c in stock_codes])
        df = pd.read_sql(
            f'SELECT * FROM financial_ratios WHERE ts_code IN ({codes_str})',
            conn
        )
        conn.close()

        if df.empty:
            print(f'  MySQL: 无匹配数据')
            return pd.DataFrame()

        # 取每只股票最近一年的数据
        df = df.sort_values('report_year', ascending=False).groupby('ts_code').first().reset_index()

        # 重命名列
        ratio_cols = {
            'current_ratio': '流动比率', 'quick_ratio': '速动比率',
            'inv_turn': '存货周转率', 'ar_turn': '应收账款周转率',
            'ca_turn': '流动资产周转率', 'assets_turn': '总资产周转率',
            'roa': 'ROA', 'npta': '总资产净利率',
            'roe': 'ROE', 'roe_dt': 'ROE摊薄',
            'cash_to_liqdebt': '现金流动负债比', 'cash_to_liqdebt_withinterest': '现金利息负债比',
            'netprofit_margin': '净利率', 'grossprofit_margin': '毛利率',
            'debt_to_assets': '资产负债率', 'int_to_talcap': '利息资本化率',
            'debt_to_eqt': '产权比率', 'ebit_to_interest': '已获利息倍数',
            'rd_exp_ratio': '研发费用率',
            'op_yoy': '营收增长', 'ebt_yoy': '利润增长',
            'netprofit_yoy': '净利增长', 'dt_netprofit_yoy': '扣非净利增长',
            'roe_yoy': 'ROE增长', 'tr_yoy': '总营收增长',
            'or_yoy': '营收增长2', 'equity_yoy': '净资产增长',
        }
        rename_map = {k: v for k, v in ratio_cols.items() if k in df.columns}
        df = df[['ts_code', 'report_year'] + list(rename_map.keys())].rename(columns=rename_map)
        df = df.rename(columns={'ts_code': '股票代码', 'report_year': '财报年份'})

        print(f'  MySQL: 匹配到 {len(df)} 家公司的财务比率')
        return df
    except Exception as e:
        print(f'  MySQL: 不可用({e})')
        return pd.DataFrame()


def main():
    parser = argparse.ArgumentParser(description='定增特征导出')
    parser.add_argument('excel_path', help='scored Excel路径')
    parser.add_argument('--output', default=None, help='输出文件路径(默认ml_training/data/features.parquet)')
    parser.add_argument('--no-mysql', action='store_true', help='跳过MySQL数据')
    args = parser.parse_args()

    # 路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_dir = os.path.dirname(os.path.dirname(script_dir))
    db_path = os.path.join(project_dir, 'price_maintenance_risk_analysis', 'data', 'valuation.db')
    output_path = args.output or os.path.join(script_dir, 'data', 'features.parquet')
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # 1. 从scored Excel获取股票列表+财务评分+标签
    print('1. 加载scored Excel...')
    scored = load_scored_features(args.excel_path)
    print(f'   {len(scored)} 只股票')

    stock_codes = scored['股票代码'].tolist()

    # 2. 从SQLite加载行情/估值/FCF特征
    print('2. 加载SQLite特征...')
    sqlite_feats = load_sqlite_features(db_path, stock_codes)
    matched = sqlite_feats['当前价'].notna().sum()
    print(f'   匹配到行情数据: {matched}/{len(stock_codes)}')

    # 3. 从MySQL加载财务比率
    if not args.no_mysql:
        print('3. 加载MySQL财务比率...')
        mysql_feats = load_mysql_features(stock_codes)
    else:
        mysql_feats = pd.DataFrame()

    # 4. 合并所有特征
    print('4. 合并特征...')
    merged = scored.merge(sqlite_feats, on='股票代码', how='left', suffixes=('', '_db'))
    if not mysql_feats.empty:
        merged = merged.merge(mysql_feats, on='股票代码', how='left')

    # 列数统计
    total_cols = len(merged.columns)
    label_cols = [c for c in merged.columns if c.startswith('标签_')]
    feature_cols = [c for c in merged.columns if not c.startswith('标签_') and c != '7个月涨跌幅']
    print(f'   总特征数: {len(feature_cols)}, 标签数: {len(label_cols)}, 总列数: {total_cols}')

    # 5. 保存
    merged.to_parquet(output_path, index=False)
    print(f'\n✅ 特征已保存: {output_path}')
    print(f'   {len(merged)} 行 × {total_cols} 列')

    # 同时保存CSV(方便人工查看)
    csv_path = output_path.replace('.parquet', '.csv')
    merged.to_csv(csv_path, index=False)
    print(f'   CSV(查看用): {csv_path}')

    # 字段说明
    schema_path = os.path.join(os.path.dirname(output_path), 'features_schema.md')
    with open(schema_path, 'w', encoding='utf-8') as f:
        f.write('# 特征字段说明\n\n')
        groups = {
            '基础信息': ['股票代码', '股票简称', '报价日', '报价日_excel', '一级行业', '二级行业', '三级行业'],
            '行情特征(SQLite)': [c for c in feature_cols if any(k in c for k in ['波动率', '收益', '胜率', 'MA', '漂移', '换手', '当前价', '数据天数'])],
            '估值特征(SQLite)': [c for c in feature_cols if any(k in c for k in ['PE', 'PB', 'PS', '行业代码', '行业名称'])],
            'FCF特征(SQLite)': [c for c in feature_cols if any(k in c for k in ['营收', 'NOPAT', 'FCF', 'FCF年份'])],
            '定增参数(SQLite)': [c for c in feature_cols if any(k in c for k in ['发行价', '锁定', '净债', '净利润', '净资产'])],
            '筛选决策(SQLite)': [c for c in feature_cols if any(k in c for k in ['溢价率', '有效阈值', 'step', '定增决策'])],
            '财务评分(Excel)': [c for c in feature_cols if any(k in c for k in ['总分', '评级', '盈利能力', '成长能力', '斜率', '趋势', '综合'])],
            '子场景(Excel)': [c for c in feature_cols if '_通过' in c or '建议参与' in c],
            '财务比率(MySQL)': [c for c in feature_cols if c in ['流动比率', '速动比率', '存货周转率', '应收账款周转率', '流动资产周转率', '总资产周转率', 'ROA', '总资产净利率', 'ROE', 'ROE摊薄', '现金流动负债比', '现金利息负债比', '净利率', '毛利率', '资产负债率', '利息资本化率', '产权比率', '已获利息倍数', '研发费用率', '营收增长', '利润增长', '净利增长', '扣非净利增长', 'ROE增长', '总营收增长', '营收增长2', '净资产增长', '财报年份']],
            '标签': label_cols + ['7个月涨跌幅', '最终结论'],
        }
        for group_name, cols in groups.items():
            if cols:
                f.write(f'\n## {group_name} ({len(cols)}个)\n\n')
                f.write('| 字段 | 来源 |\n|------|------|\n')
                for c in cols:
                    src = 'SQLite' if c in sqlite_feats.columns else ('MySQL' if c in mysql_feats.columns else 'Excel')
                    f.write(f'| {c} | {src} |\n')
    print(f'   字段说明: {schema_path}')


if __name__ == '__main__':
    main()
