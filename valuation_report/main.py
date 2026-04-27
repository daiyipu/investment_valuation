#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
上市公司估值分析报告生成器

使用方法：
    python main.py --stock 002276.SZ --name 万马高分子
    python main.py --stock 600519.SH
    python main.py --stock 300750.SZ --chapter 4
"""

import sys
import os
import yaml
import time
import argparse
from datetime import datetime, timedelta

import tushare as ts
import pandas as pd
from docx import Document

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

import module_utils as utils
from utils.tushare_client import TushareClient
from utils.font_manager import get_font_prop
from utils.wacc_calculator import WACCCalculator

# 导入章节模块
from chapters import (
    chapter01_overview, chapter02_financial, chapter03_industry,
    chapter04_dcf, chapter05_relative, chapter06_sensitivity,
    chapter07_montecarlo, chapter08_scenario, chapter09_stress,
    chapter10_var, chapter11_assessment,
)
# Ensure each module has generate_chapter imported
for _mod in [chapter01_overview, chapter02_financial, chapter03_industry,
             chapter04_dcf, chapter05_relative, chapter06_sensitivity,
             chapter07_montecarlo, chapter08_scenario, chapter09_stress,
             chapter10_var, chapter11_assessment]:
    if not hasattr(_mod, 'generate_chapter'):
        from importlib import import_module
        _mod.generate_chapter = getattr(import_module(_mod.__name__), 'generate_chapter', None)

IMAGES_DIR = os.path.join(SCRIPT_DIR, 'output', 'images')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')
os.makedirs(IMAGES_DIR, exist_ok=True)


def load_config():
    config_path = os.path.join(SCRIPT_DIR, 'config.yaml')
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
    else:
        config = {}
    # 环境变量优先
    token = os.environ.get('TUSHARE_TOKEN', '')
    if token:
        config.setdefault('tushare', {})['token'] = token
    return config


def generate_report(stock_code, stock_name='', target_chapter=None):
    config = load_config()
    token = config.get('tushare', {}).get('token', '')
    if not token:
        print("错误: 未设置TUSHARE_TOKEN环境变量或config.yaml中的token")
        sys.exit(1)

    print("=" * 70)
    print(" 上市公司估值分析报告生成器")
    print("=" * 70)

    # 初始化Tushare
    tc = TushareClient(token)
    pro = tc.get_pro()
    font_prop = get_font_prop()

    # 自动获取股票名称
    if not stock_name:
        df_basic = pro.stock_basic(ts_code=stock_code, fields='ts_code,name')
        if not df_basic.empty:
            stock_name = df_basic.iloc[0]['name']
        else:
            stock_name = stock_code
    report_date = datetime.now().strftime('%Y-%m-%d')
    print(f"股票代码: {stock_code}")
    print(f"股票名称: {stock_name}")
    print(f"报告日期: {report_date}")
    print("=" * 70)

    # ============ 数据加载 ============
    print("\n[1/8] 加载公司基本信息...")
    company_info = _load_company_info(pro, stock_code)
    if company_info:
        print(f"  公司: {company_info.get('name', stock_name)} 行业: {company_info.get('industry', 'N/A')}")

    print("[2/8] 查询SW行业分类与同行...")
    sw_industry, peer_codes = _load_sw_industry(pro, stock_code)
    if sw_industry:
        print(f"  SW三级: {sw_industry.get('l3_name', 'N/A')} 同行: {len(peer_codes)}家")

    print("[3/8] 获取行情数据...")
    price_data = _load_price_data(pro, stock_code)
    daily_basic = _load_daily_basic(pro, stock_code)
    if daily_basic:
        print(f"  最新PE={daily_basic.get('pe_ttm', 'N/A')} PB={daily_basic.get('pb', 'N/A')} 市值={daily_basic.get('total_mv', 'N/A')}万")

    print("[4/8] 获取财务报表(5年)...")
    financial_statements = _load_financial_statements(pro, stock_code)
    financial_indicators = _load_financial_indicators(pro, stock_code)
    print(f"  利润表: {len(financial_statements.get('income', pd.DataFrame()))}期 "
          f"资产负债表: {len(financial_statements.get('balancesheet', pd.DataFrame()))}期 "
          f"现金流量表: {len(financial_statements.get('cashflow', pd.DataFrame()))}期")

    print("[5/8] 获取历史估值(3年)...")
    historical_valuation = _load_historical_valuation(pro, stock_code)
    print(f"  数据点: {len(historical_valuation)}个交易日")

    print("[6/8] 获取同行估值数据...")
    peer_valuation = _load_peer_valuation(pro, peer_codes, daily_basic)
    print(f"  有效同行: {len(peer_valuation)}家")

    print("[7/8] 计算WACC...")
    wacc_result = _calculate_wacc(pro, stock_code, config, daily_basic, company_info)
    _wacc_val = wacc_result.get('wacc', 0)
    if isinstance(_wacc_val, dict):
        _wacc_val = _wacc_val.get('wacc', 0)
    print(f"  WACC={_wacc_val*100:.2f}% Beta={wacc_result.get('beta', 0):.3f}")

    print("[8/8] 获取市场指数...")
    market_indices = _load_market_indices(pro)
    print(f"  指数: {list(market_indices.keys())}")

    # ============ 创建文档 ============
    document = Document()
    utils.setup_chinese_font(document)

    context = {
        'stock_code': stock_code,
        'stock_name': stock_name,
        'report_date': report_date,
        'document': document,
        'font_prop': font_prop,
        'config': config,
        'pro': pro,

        'company_info': company_info,
        'sw_industry': sw_industry,
        'peer_codes': peer_codes,
        'price_data': price_data,
        'daily_basic': daily_basic,
        'market_indices': market_indices,
        'financial_statements': financial_statements,
        'financial_indicators': financial_indicators,
        'historical_valuation': historical_valuation,
        'peer_valuation': peer_valuation,
        'wacc_result': wacc_result,

        'IMAGES_DIR': IMAGES_DIR,
        'OUTPUT_DIR': OUTPUT_DIR,
        'results': {},
    }

    # ============ 章节编排 ============
    chapters = [
        (1, '项目概述', chapter01_overview),
        (2, '财务分析', chapter02_financial),
        (3, '行业分析', chapter03_industry),
        (4, 'DCF绝对估值', chapter04_dcf),
        (5, '相对估值', chapter05_relative),
        (6, '敏感性分析', chapter06_sensitivity),
        (7, '蒙特卡洛模拟', chapter07_montecarlo),
        (8, '情景分析', chapter08_scenario),
        (9, '压力测试', chapter09_stress),
        (10, 'VaR风险度量', chapter10_var),
        (11, '综合评估', chapter11_assessment),
    ]

    for ch_num, ch_name, ch_module in chapters:
        if target_chapter and ch_num != target_chapter:
            continue
        print(f"\n生成第{ch_num}章：{ch_name}...")
        try:
            context = ch_module.generate_chapter(context)
        except Exception as e:
            print(f"  第{ch_num}章生成失败: {e}")
            import traceback
            traceback.print_exc()

    # 页眉页码
    utils.setup_document_header(document, stock_name, stock_code)
    utils.add_page_numbers(document)

    # 保存
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{stock_name}_{stock_code}_估值分析报告_{timestamp}.docx"
    output_path = os.path.join(OUTPUT_DIR, filename)
    document.save(output_path)
    print(f"\n报告已保存: {output_path}")
    return output_path


# ==================== 数据加载函数 ====================

def _load_company_info(pro, stock_code):
    try:
        df = pro.stock_basic(ts_code=stock_code, fields='ts_code,name,industry,market,list_date')
        if df.empty:
            return {}
        row = df.iloc[0]
        df_company = pro.stock_basic(ts_code=stock_code, fields='ts_code,name,industry,market,list_date,enname,chairman,province,city,website')
        info = df_company.iloc[0].to_dict() if not df_company.empty else row.to_dict()
        return info
    except Exception as e:
        print(f"  加载公司信息失败: {e}")
        return {}


def _load_sw_industry(pro, stock_code):
    sw_industry = {}
    peer_codes = []
    try:
        df = pro.index_member_all(ts_code=stock_code)
        if df is None or df.empty:
            return sw_industry, peer_codes
        df = df[df['ts_code'] == stock_code].sort_values('in_date', ascending=False)
        if df.empty:
            return sw_industry, peer_codes
        latest = df.iloc[0]
        l1_code = latest.get('l1_code', '')
        l2_code = latest.get('l2_code', '')
        l3_code = latest.get('l3_code', '')
        sw_industry = {
            'l1_name': latest.get('l1_name', ''),
            'l1_code': l1_code,
            'l2_name': latest.get('l2_name', ''),
            'l2_code': l2_code,
            'l3_name': latest.get('l3_name', ''),
            'l3_code': l3_code,
        }

        # Step 1: Try L3 via index_member_all
        if l3_code:
            df_peers = pro.index_member_all(l3_code=l3_code)
            if df_peers is not None and not df_peers.empty:
                peer_codes = [c for c in df_peers['ts_code'].unique().tolist() if c != stock_code]

        # Step 2: If L3 too few, try sibling L3 industries (same parent L2)
        if len(peer_codes) < 5 and l2_code:
            print(f"  SW三级同行仅{len(peer_codes)}家，尝试同级兄弟行业")
            df_l3_list = pro.index_classify(level='L3', src='SW2021')
            if df_l3_list is not None and not df_l3_list.empty:
                # Find all L3 under the same L2 parent
                siblings = df_l3_list[df_l3_list['parent_code'] == l2_code]
                for _, sib in siblings.iterrows():
                    sib_code = sib['index_code']
                    if sib_code == l3_code:
                        continue
                    try:
                        sib_members = pro.index_member(index_code=sib_code, is_new='Y')
                        if sib_members is not None and not sib_members.empty:
                            for code in sib_members['con_code'].unique().tolist():
                                if code != stock_code and code not in peer_codes:
                                    peer_codes.append(code)
                        time.sleep(0.2)
                    except Exception:
                        pass
            if len(peer_codes) >= 5:
                print(f"  合并兄弟行业后同行{len(peer_codes)}家")

        # Step 3: If still too few, try L2 via index_member
        if len(peer_codes) < 5 and l2_code:
            print(f"  尝试SW二级行业成员")
            try:
                l2_members = pro.index_member(index_code=l2_code, is_new='Y')
                if l2_members is not None and not l2_members.empty:
                    for code in l2_members['con_code'].unique().tolist():
                        if code != stock_code and code not in peer_codes:
                            peer_codes.append(code)
                print(f"  扩展到L2后同行{len(peer_codes)}家")
            except Exception:
                pass

        if len(peer_codes) < 5 and peer_codes:
            print(f"  注意: 同行仅{len(peer_codes)}家，统计意义有限")
    except Exception as e:
        print(f"  加载SW行业失败: {e}")
    return sw_industry, peer_codes


def _load_price_data(pro, stock_code, months=3):
    try:
        end = datetime.now().strftime('%Y%m%d')
        start = (datetime.now() - timedelta(days=months * 31)).strftime('%Y%m%d')
        df = pro.daily(ts_code=stock_code, start_date=start, end_date=end)
        if df is not None and not df.empty:
            df = df.sort_values('trade_date').reset_index(drop=True)
        return df if df is not None else pd.DataFrame()
    except Exception as e:
        print(f"  加载行情失败: {e}")
        return pd.DataFrame()


def _load_daily_basic(pro, stock_code):
    try:
        for days_back in range(1, 8):
            test_date = (datetime.now() - timedelta(days=days_back)).strftime('%Y%m%d')
            df = pro.daily_basic(ts_code=stock_code, trade_date=test_date,
                                 fields='ts_code,trade_date,close,pe_ttm,pb,ps_ttm,total_mv,total_share')
            if df is not None and not df.empty:
                return df.iloc[0].to_dict()
        return {}
    except Exception as e:
        print(f"  加载每日指标失败: {e}")
        return {}


def _load_financial_statements(pro, stock_code, years=5):
    result = {'income': pd.DataFrame(), 'balancesheet': pd.DataFrame(), 'cashflow': pd.DataFrame()}
    try:
        start_year = datetime.now().year - years
        start_date = f"{start_year}0101"
        for stmt_name, api_func in [('income', pro.income), ('balancesheet', pro.balancesheet), ('cashflow', pro.cashflow)]:
            df = api_func(ts_code=stock_code, start_date=start_date)
            if df is not None and not df.empty:
                df = df.sort_values('end_date', ascending=False).reset_index(drop=True)
                result[stmt_name] = df
                time.sleep(0.3)
    except Exception as e:
        print(f"  加载财务报表失败: {e}")
    return result


def _load_financial_indicators(pro, stock_code, years=5):
    try:
        start_year = datetime.now().year - years
        start_date = f"{start_year}0101"
        df = pro.fina_indicator(ts_code=stock_code, start_date=start_date)
        if df is not None and not df.empty:
            return df.sort_values('end_date', ascending=False).reset_index(drop=True)
    except Exception as e:
        print(f"  加载财务指标失败: {e}")
    return pd.DataFrame()


def _load_historical_valuation(pro, stock_code, years=3):
    try:
        end = datetime.now().strftime('%Y%m%d')
        start = (datetime.now() - timedelta(days=years * 366)).strftime('%Y%m%d')
        all_data = []
        current_date = start
        while current_date < end:
            df = pro.daily_basic(ts_code=stock_code, start_date=current_date,
                                 end_date=min((pd.Timestamp(current_date) + timedelta(days=200)).strftime('%Y%m%d'), end),
                                 fields='ts_code,trade_date,pe_ttm,pb,ps_ttm,close')
            if df is not None and not df.empty:
                all_data.append(df)
            current_date = (pd.Timestamp(current_date) + timedelta(days=200)).strftime('%Y%m%d')
            time.sleep(0.3)
        if all_data:
            valid = [df for df in all_data if not df.empty and not df.isna().all(axis=None)]
            if valid:
                return pd.concat(valid, ignore_index=True).sort_values('trade_date').reset_index(drop=True)
    except Exception as e:
        print(f"  加载历史估值失败: {e}")
    return pd.DataFrame()


def _load_peer_valuation(pro, peer_codes, daily_basic):
    if not peer_codes:
        return pd.DataFrame()
    trade_date = daily_basic.get('trade_date', '')
    if not trade_date:
        return pd.DataFrame()
    results = []
    for code in peer_codes[:30]:
        try:
            df = pro.daily_basic(ts_code=code, trade_date=trade_date,
                                 fields='ts_code,pe_ttm,pb,ps_ttm,total_mv')
            if df is not None and not df.empty:
                df_basic = pro.stock_basic(ts_code=code, fields='ts_code,name')
                name = df_basic.iloc[0]['name'] if not df_basic.empty else code
                row = df.iloc[0].to_dict()
                row['name'] = name
                results.append(row)
            time.sleep(0.2)
        except Exception:
            pass
    if results:
        return pd.DataFrame(results)
    return pd.DataFrame()


def _calculate_wacc(pro, stock_code, config, daily_basic, company_info):
    wacc_config = config.get('wacc', {})
    # 传入已有市值（万元→元），避免get_capital_structure重新查询失败
    existing_market_cap = float(daily_basic.get('total_mv', 0)) * 10000 if daily_basic.get('total_mv') else 0
    try:
        calculator = WACCCalculator(pro=pro, params=wacc_config)
        result = calculator.calculate_wacc(
            stock_code=stock_code,
            use_industry_beta=True,
        )
        # 如果股权占比异常（为0或100%），用已有市值重算
        equity_ratio = result.get('equity_ratio', 0)
        if existing_market_cap > 0 and (equity_ratio < 0.05 or equity_ratio > 0.99):
            total_debt = result.get('total_debt', 0)
            total_capital = existing_market_cap + total_debt
            if total_capital > 0:
                result['equity_ratio'] = existing_market_cap / total_capital
                result['debt_ratio'] = total_debt / total_capital
                result['market_cap'] = existing_market_cap
                ke = result.get('ke', result.get('cost_of_equity', 0))
                kd_aftertax = result.get('kd_aftertax', result.get('cost_of_debt', 0) * (1 - 0.25))
                result['wacc'] = result['equity_ratio'] * ke + result['debt_ratio'] * kd_aftertax
                print(f"  已用市值数据修正资本结构: 股权{result['equity_ratio']:.1%} 债务{result['debt_ratio']:.1%}")
        # 统一字段名
        result.setdefault('cost_of_equity', result.get('ke', 0))
        result.setdefault('cost_of_debt', result.get('kd_pretax', 0))
        result.setdefault('beta', result.get('adopted_beta', result.get('stock_beta', 1.0)))
        result.setdefault('capital_structure', {
            'equity_ratio': result.get('equity_ratio', 0.7),
            'debt_ratio': result.get('debt_ratio', 0.3),
        })
        result.setdefault('parameters', wacc_config)
        return result
    except Exception as e:
        print(f"  WACC计算失败({e})，使用简化计算")
        from utils.wacc_calculator import calculate_wacc_simple
        beta = wacc_config.get('fallback_beta', 1.0)
        simple = calculate_wacc_simple(beta=beta)
        return {
            'wacc': simple['wacc'],
            'beta': beta,
            'cost_of_equity': simple['ke'],
            'cost_of_debt': simple['kd_pretax'],
            'capital_structure': {'equity_ratio': 0.7, 'debt_ratio': 0.3},
            'parameters': wacc_config,
        }


def _load_market_indices(pro):
    indices = {}
    for name, ts_code in [('沪深300', '000300.SH'), ('上证指数', '000001.SH'), ('深证成指', '399001.SZ')]:
        try:
            end = datetime.now().strftime('%Y%m%d')
            start = (datetime.now() - timedelta(days=365)).strftime('%Y%m%d')
            df = pro.index_daily(ts_code=ts_code, start_date=start, end_date=end)
            if df is not None and not df.empty:
                df = df.sort_values('trade_date').reset_index(drop=True)
                indices[name] = df
            time.sleep(0.2)
        except Exception:
            pass
    return indices


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='上市公司估值分析报告生成器')
    parser.add_argument('--stock', type=str, required=True, help='股票代码（如 002276.SZ）')
    parser.add_argument('--name', type=str, default='', help='股票名称')
    parser.add_argument('--chapter', type=int, default=None, help='仅生成指定章节（调试用）')
    args = parser.parse_args()

    stock_code = args.stock.strip().upper()
    # Auto-correct suffix: 6xx→.SH, 0xx/3xx→.SZ
    if '.' not in stock_code:
        stock_code = stock_code + ('.SH' if stock_code.startswith('6') else '.SZ')
    else:
        pure = stock_code.split('.')[0]
        expected = '.SH' if pure.startswith('6') else '.SZ'
        if not stock_code.endswith(expected):
            corrected = pure + expected
            print(f"注意: 股票代码后缀已自动修正 {stock_code} → {corrected}")
            stock_code = corrected

    generate_report(stock_code, args.name, args.chapter)
