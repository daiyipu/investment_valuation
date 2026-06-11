#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增特征衍生脚本 - 从基础特征派生高级特征

两阶段流水线:
  Stage 1 (A-E): 纯 parquet 运算，不访问数据库
  Stage 2 (F-G): 需 MySQL 查询 (industry_daily, market_indices)

用法:
    python ml_training/derive_features.py [input.parquet] [--output features_derived.parquet] [--no-db]

输出:
    ml_training/data/features_derived.parquet - 原始特征 + 衍生特征
"""

import sys
import os
import argparse
import numpy as np
import pandas as pd


# ====== 通用工具 ======

def _safe_divide(a, b, fillna=np.nan):
    """安全除法: abs(b)<1e-8 返回 NaN"""
    a = pd.to_numeric(a, errors='coerce')
    b = pd.to_numeric(b, errors='coerce')
    result = np.where(np.abs(b) < 1e-8, fillna, a / b)
    return pd.Series(result, index=a.index)


def _safe_yoy(current, prior):
    """YoY增长率: (current - prior) / |prior|"""
    current = pd.to_numeric(current, errors='coerce')
    prior = pd.to_numeric(prior, errors='coerce')
    return _safe_divide(current - prior, prior)


def _check_cols(df, required, category):
    """检查必需列是否存在，缺失则返回False"""
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f'  ⚠️ {category}: 缺少列 {missing[:5]}{"..." if len(missing)>5 else ""}，跳过')
        return False
    return True


# ====== A类: FCF 增长率 (+18特征) ======

FCF_METRICS = ['营收', 'NOPAT', 'FCF', '营业利润', '净利润_fcf', '资本支出']


def derive_fcf_growth_rates(df):
    """从 _T/_T1/_T2 列计算 YoY、CAGR2、增长加速度"""
    print('\n  A类: FCF增长率...')
    required = []
    for m in FCF_METRICS:
        for s in ['_T', '_T1', '_T2']:
            required.append(f'{m}{s}')
    if not _check_cols(df, required, 'A类-FCF增长率'):
        return df

    new_cols = {}
    for m in FCF_METRICS:
        t  = pd.to_numeric(df[f'{m}_T'],  errors='coerce')
        t1 = pd.to_numeric(df[f'{m}_T1'], errors='coerce')
        t2 = pd.to_numeric(df[f'{m}_T2'], errors='coerce')

        # YoY增长率
        new_cols[f'{m}_YoY'] = _safe_yoy(t, t1)

        # 2年CAGR: (T/T2)^0.5 - 1 (仅当同号时)
        ratio = _safe_divide(t, t2)
        same_sign = (t * t2) > 0
        cagr = np.where(same_sign, np.power(ratio.abs(), 0.5) - 1, np.nan)
        cagr = np.where(same_sign & (t2 > 0), ratio.pow(0.5) - 1, cagr)
        new_cols[f'{m}_CAGR2'] = pd.Series(cagr, index=df.index)

        # 增长加速度: YoY_T - YoY_T1
        yoy_t  = _safe_yoy(t, t1)
        yoy_t1 = _safe_yoy(t1, t2)
        new_cols[f'{m}_加速'] = yoy_t - yoy_t1

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in list(coverage.items())[:6])}...')
    return df


# ====== B类: FCF 交叉比率 (+5特征) ======

def derive_fcf_cross_metrics(df):
    """FCF相关的财务比率"""
    print('\n  B类: FCF交叉比率...')
    required = ['FCF_T', '营收_T', 'NOPAT_T', '资本支出_T', '折旧_T']
    if not _check_cols(df, required, 'B类-FCF交叉比率'):
        return df

    fcf    = pd.to_numeric(df['FCF_T'],    errors='coerce')
    rev    = pd.to_numeric(df['营收_T'],   errors='coerce')
    nopat  = pd.to_numeric(df['NOPAT_T'],  errors='coerce')
    capex  = pd.to_numeric(df['资本支出_T'], errors='coerce')
    dep    = pd.to_numeric(df['折旧_T'],   errors='coerce')

    new_cols = {
        'FCF_margin':     _safe_divide(fcf, rev),
        'FCF_conversion': _safe_divide(fcf, nopat),
        'capex_to_dep':   _safe_divide(capex, dep),
        'capex_intensity': _safe_divide(capex, rev),
        'NOPAT_margin':   _safe_divide(nopat, rev),
    }
    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


# ====== C类: 财务评分变动 (+9特征) ======

SCORE_METRICS = ['总分', '盈利能力', '成长能力']
SCORE_DELTAS = [('delta_1y', 'T', 'T-1'), ('delta_2y', 'T', 'T-2'), ('delta_4y', 'T', 'T-4')]


def derive_financial_score_deltas(df):
    """财务评分的年度变动量"""
    print('\n  C类: 财务评分变动...')
    # 检查必需列: {metric}_T 和 {metric}_T-1 等
    required = set()
    for m in SCORE_METRICS:
        for _, y_new, y_old in SCORE_DELTAS:
            required.add(f'{m}_{y_new}')
            required.add(f'{m}_{y_old}')
    if not _check_cols(df, list(required), 'C类-评分变动'):
        return df

    new_cols = {}
    for m in SCORE_METRICS:
        for suffix, y_new, y_old in SCORE_DELTAS:
            col_new = f'{m}_{y_new}'
            col_old = f'{m}_{y_old}'
            if col_new in df.columns and col_old in df.columns:
                v_new = pd.to_numeric(df[col_new], errors='coerce')
                v_old = pd.to_numeric(df[col_old], errors='coerce')
                new_cols[f'{m}_{suffix}'] = v_new - v_old

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in list(coverage.items())[:6])}...')
    return df


# ====== D类: 估值相对特征 (+7特征) ======

def derive_valuation_relative(df):
    """个股估值 vs 行业/同行"""
    print('\n  D类: 估值相对特征...')

    new_cols = {}

    # vs 行业
    pairs_industry = [('PE', '个股PE', '行业PE'), ('PB', '个股PB', '行业PB')]
    for val_type, col_stock, col_ind in pairs_industry:
        if col_stock in df.columns and col_ind in df.columns:
            new_cols[f'{val_type}_vs_行业'] = _safe_divide(
                pd.to_numeric(df[col_stock], errors='coerce'),
                pd.to_numeric(df[col_ind], errors='coerce')
            )

    # vs 同行
    pairs_peer = [
        ('PE', '个股PE', '同行PE_均值', 'PE_vs_同行均值'),
        ('PE', '个股PE', '同行PE_中位', 'PE_vs_同行中位'),
        ('PB', '个股PB', '同行PB_均值', 'PB_vs_同行均值'),
        ('PB', '个股PB', '同行PB_中位', 'PB_vs_同行中位'),
        ('PS', '个股PS', '同行PS_均值', 'PS_vs_同行均值'),
    ]
    for val_type, col_stock, col_peer, col_name in pairs_peer:
        if col_stock in df.columns and col_peer in df.columns:
            new_cols[col_name] = _safe_divide(
                pd.to_numeric(df[col_stock], errors='coerce'),
                pd.to_numeric(df[col_peer], errors='coerce')
            )

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


# ====== E类: 行情动量与位置 (+7特征) ======

def derive_market_momentum(df):
    """波动率比值、动量加速、价格vs均线"""
    print('\n  E类: 行情动量与位置...')

    new_cols = {}

    # 波动率比值
    for col_name, w_short, w_long in [('vol_ratio_20_60', '波动率_20d', '波动率_60d'),
                                       ('vol_ratio_60_250', '波动率_60d', '波动率_250d')]:
        if w_short in df.columns and w_long in df.columns:
            new_cols[col_name] = _safe_divide(
                pd.to_numeric(df[w_short], errors='coerce'),
                pd.to_numeric(df[w_long], errors='coerce')
            )

    # 动量加速
    if '年化收益_60d' in df.columns and '年化收益_120d' in df.columns:
        new_cols['return_acceleration'] = (
            pd.to_numeric(df['年化收益_60d'], errors='coerce') -
            pd.to_numeric(df['年化收益_120d'], errors='coerce')
        )

    # 价格 vs 均线
    if '当前价' in df.columns:
        price = pd.to_numeric(df['当前价'], errors='coerce')
        for ma_label in ['MA20', 'MA60', 'MA120', 'MA250']:
            if ma_label in df.columns:
                new_cols[f'price_vs_{ma_label}'] = _safe_divide(price, pd.to_numeric(df[ma_label], errors='coerce'))

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in list(coverage.items())[:6])}...')
    return df


# ====== F类: 行业PE/PB增长率 (+6特征, 需DB) ======

def derive_industry_valuation_growth(df):
    """从 industry_daily 表计算行业PE/PB在报价日前的增长率"""
    print('\n  F类: 行业PE/PB增长率 (需MySQL)...')

    import pymysql
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')

    # 从 industry_data 表获取每只股票对应的 index_code（与 industry_daily 一致）
    stock_codes = df['股票代码'].unique().tolist()
    codes_ph = ','.join(['%s'] * len(stock_codes))
    id_df = pd.read_sql(
        f'SELECT stock_code, index_code FROM industry_data WHERE stock_code IN ({codes_ph})',
        conn, params=stock_codes
    )
    # 建立 stock_code → index_code 映射
    stock_to_index = dict(zip(id_df['stock_code'], id_df['index_code']))
    print(f'    行业映射: {len(stock_to_index)} 只股票')

    # 批量加载所有 industry_daily 数据
    daily_df = pd.read_sql(
        'SELECT index_code, trade_date, pe, pb, close FROM industry_daily ORDER BY index_code, trade_date',
        conn
    )
    conn.close()

    if daily_df.empty:
        print('    ⚠️ industry_daily 无数据，跳过')
        return df

    print(f'    加载行业日线: {daily_df["index_code"].nunique()}个行业, {len(daily_df)}条')

    # 按 index_code 分组
    daily_groups = {}
    for code, group in daily_df.groupby('index_code'):
        g = group.sort_values('trade_date').reset_index(drop=True)
        daily_groups[code] = {
            'dates': g['trade_date'].values,
            'pe': g['pe'].values.astype(float),
            'pb': g['pb'].values.astype(float),
        }

    # 获取每条样本的 index_code
    sample_index_codes = df['股票代码'].map(stock_to_index)

    # 对每条样本计算
    growth_windows = [60, 120, 250]
    pe_growth = {w: np.full(len(df), np.nan) for w in growth_windows}
    pb_growth = {w: np.full(len(df), np.nan) for w in growth_windows}

    matched = 0
    for idx, row in df.iterrows():
        code = sample_index_codes.iloc[idx] if idx < len(sample_index_codes) else None
        if pd.isna(code) or code not in daily_groups:
            continue
        issue_date_raw = row.get('报价日')
        if pd.isna(issue_date_raw):
            continue
        issue_str = str(int(float(issue_date_raw))) if isinstance(issue_date_raw, (int, float)) else str(issue_date_raw)
        if len(issue_str) < 8:
            continue

        dg = daily_groups[code]
        # 找到 <= issue_str 的最大索引
        mask = dg['dates'] <= issue_str
        if not mask.any():
            continue
        pos = mask.sum() - 1  # 最后一个 <= issue_str 的位置

        pe_now = dg['pe'][pos]
        pb_now = dg['pb'][pos]
        if np.isnan(pe_now) and np.isnan(pb_now):
            continue

        for w in growth_windows:
            if pos >= w:
                pe_prev = dg['pe'][pos - w]
                pb_prev = dg['pb'][pos - w]
                if not np.isnan(pe_now) and not np.isnan(pe_prev) and abs(pe_prev) > 1e-8:
                    pe_growth[w][idx] = pe_now / pe_prev - 1
                if not np.isnan(pb_now) and not np.isnan(pb_prev) and abs(pb_prev) > 1e-8:
                    pb_growth[w][idx] = pb_now / pb_prev - 1
        matched += 1

    print(f'    匹配行业日线: {matched}/{len(df)}')

    new_cols = {}
    for w in growth_windows:
        new_cols[f'行业PE_{w}d增长'] = pd.Series(pe_growth[w], index=df.index)
        new_cols[f'行业PB_{w}d增长'] = pd.Series(pb_growth[w], index=df.index)

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


# ====== G类: 大盘指数特征 (+6特征, 需DB) ======

# 交易所 → 指数名 映射
EXCHANGE_INDEX_MAP = {
    '300': '创业板指',
    '301': '创业板指',
    '688': '科创50',
    '60':  '沪深300',
    '000': '中证500',
    '001': '中证500',
    '002': '中证500',
    '003': '中证500',
}
FALLBACK_INDEX = '沪深300'


def _get_index_name(stock_code):
    """根据股票代码前缀判断对应的市场指数"""
    for prefix, idx_name in sorted(EXCHANGE_INDEX_MAP.items(), key=lambda x: -len(x[0])):
        if stock_code.startswith(prefix):
            return idx_name
    return FALLBACK_INDEX


def derive_market_index_features(df):
    """从 market_indices 表获取报价日时的大盘特征"""
    print('\n  G类: 大盘指数特征 (需MySQL)...')

    import pymysql
    conn = pymysql.connect(host='127.0.0.1', port=3306, user='root', password='',
                           database='investment_valuation', charset='utf8mb4')

    # 批量加载所有数据
    query = '''SELECT index_name, locked_date, current_level,
                      volatility_60d, volatility_120d, volatility_250d,
                      return_60d, return_120d,
                      ma_250, win_rate_120d
               FROM market_indices'''
    mi_df = pd.read_sql(query, conn)
    conn.close()

    if mi_df.empty:
        print('    ⚠️ market_indices 无数据，跳过')
        return df

    print(f'    加载大盘指数: {mi_df["index_name"].nunique()}个指数, {len(mi_df)}条')

    # 构建 (index_name, locked_date) → row 的查找字典
    mi_lookup = {}
    for _, row in mi_df.iterrows():
        key = (row['index_name'], str(row['locked_date']))
        mi_lookup[key] = row

    # 对每条样本查找
    mkt_vol120 = np.full(len(df), np.nan)
    mkt_ret120 = np.full(len(df), np.nan)
    mkt_wr120  = np.full(len(df), np.nan)
    mkt_above_ma250 = np.full(len(df), np.nan)
    mkt_dist_ma250  = np.full(len(df), np.nan)
    mkt_vol_ratio   = np.full(len(df), np.nan)

    for idx, row in df.iterrows():
        stock_code = str(row.get('股票代码', ''))
        issue_date_raw = row.get('报价日')
        if pd.isna(issue_date_raw):
            continue
        issue_str = str(int(float(issue_date_raw))) if isinstance(issue_date_raw, (int, float)) else str(issue_date_raw)
        if len(issue_str) < 8:
            continue

        index_name = _get_index_name(stock_code)

        # 尝试精确匹配，找不到则往前找最近的
        mi_row = None
        for key in [(index_name, issue_str)]:
            if key in mi_lookup:
                mi_row = mi_lookup[key]
                break

        if mi_row is None:
            # 查找该指数在该日期之前的最近日期
            candidates = mi_df[(mi_df['index_name'] == index_name) &
                              (mi_df['locked_date'] <= issue_str)]
            if candidates.empty:
                # 回退到沪深300
                candidates = mi_df[(mi_df['index_name'] == FALLBACK_INDEX) &
                                  (mi_df['locked_date'] <= issue_str)]
            if candidates.empty:
                continue
            mi_row = candidates.iloc[-1]

        level = mi_row.get('current_level')
        vol120 = mi_row.get('volatility_120d')
        vol60  = mi_row.get('volatility_60d')
        vol250 = mi_row.get('volatility_250d')
        ret120 = mi_row.get('return_120d')
        wr120  = mi_row.get('win_rate_120d')
        ma250  = mi_row.get('ma_250')

        if pd.notna(vol120):
            mkt_vol120[idx] = float(vol120)
        if pd.notna(ret120):
            mkt_ret120[idx] = float(ret120)
        if pd.notna(wr120):
            mkt_wr120[idx] = float(wr120)
        if pd.notna(level) and pd.notna(ma250) and float(ma250) > 0:
            mkt_above_ma250[idx] = 1.0 if float(level) > float(ma250) else 0.0
            mkt_dist_ma250[idx] = (float(level) - float(ma250)) / float(ma250)
        if pd.notna(vol60) and pd.notna(vol250) and abs(float(vol250)) > 1e-8:
            mkt_vol_ratio[idx] = float(vol60) / float(vol250)

    new_cols = {
        '市场波动率_120d':   pd.Series(mkt_vol120, index=df.index),
        '市场年化收益_120d': pd.Series(mkt_ret120, index=df.index),
        '市场胜率_120d':     pd.Series(mkt_wr120, index=df.index),
        '市场_above_MA250':  pd.Series(mkt_above_ma250, index=df.index),
        '市场距离MA250':     pd.Series(mkt_dist_ma250, index=df.index),
        '市场波动率比值':    pd.Series(mkt_vol_ratio, index=df.index),
    }

    for k, v in new_cols.items():
        df[k] = v
    coverage = {k: f'{v.notna().mean()*100:.1f}%' for k, v in new_cols.items()}
    print(f'    +{len(new_cols)}个特征: {", ".join(f"{k}({v})" for k,v in coverage.items())}')
    return df


# ====== 主流程 ======

def main():
    parser = argparse.ArgumentParser(description='定增特征衍生 - 从基础特征派生高级特征')
    parser.add_argument('input_path', nargs='?',
                        default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'features.parquet'),
                        help='输入特征文件 (parquet或CSV)')
    parser.add_argument('--output', default=None,
                        help='输出路径 (默认: 同目录下 features_derived.parquet)')
    parser.add_argument('--no-db', action='store_true', help='跳过需MySQL的特征(F+G类)')
    args = parser.parse_args()

    # 输出路径
    if args.output:
        output_path = args.output
    else:
        base_dir = os.path.dirname(args.input_path)
        output_path = os.path.join(base_dir, 'features_derived.parquet')

    # 加载
    print(f'加载基础特征: {args.input_path}')
    if args.input_path.endswith('.parquet'):
        df = pd.read_parquet(args.input_path)
    else:
        df = pd.read_csv(args.input_path)
    orig_cols = len(df.columns)
    print(f'  {len(df)} 行 × {orig_cols} 列')

    # ====== Stage 1: Parquet-only ======
    print('\n' + '='*60)
    print('Stage 1: Parquet-only 衍生特征')
    print('='*60)

    df = derive_fcf_growth_rates(df)
    df = derive_fcf_cross_metrics(df)
    df = derive_financial_score_deltas(df)
    df = derive_valuation_relative(df)
    df = derive_market_momentum(df)

    # ====== Stage 2: DB-dependent ======
    if not args.no_db:
        print('\n' + '='*60)
        print('Stage 2: DB-dependent 衍生特征')
        print('='*60)
        try:
            df = derive_industry_valuation_growth(df)
            df = derive_market_index_features(df)
        except Exception as e:
            print(f'  ⚠️ DB特征失败: {e}')

    # ====== 清理 ======
    df = df.replace([np.inf, -np.inf], np.nan)

    # ====== 保存 ======
    new_cols_count = len(df.columns) - orig_cols
    df.to_parquet(output_path, index=False)
    csv_path = output_path.replace('.parquet', '.csv')
    df.to_csv(csv_path, index=False)

    print('\n' + '='*60)
    print(f'✅ 衍生特征已保存: {output_path}')
    print(f'   {len(df)} 行 × {len(df.columns)} 列 (原始{orig_cols} + 新增{new_cols_count})')
    print(f'   CSV(查看用): {csv_path}')

    # 新增特征覆盖率汇总
    new_col_names = df.columns[orig_cols:].tolist()
    if new_col_names:
        print(f'\n新增特征覆盖率:')
        for c in new_col_names:
            rate = df[c].notna().mean() * 100
            print(f'  {c:30s} {rate:5.1f}%')


if __name__ == '__main__':
    main()
