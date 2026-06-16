#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
定增风险分析报告生成器 - 统一入口（模块化版本）

功能：
    统一调用各章节模块，生成完整的Word分析报告

使用方法：
    python -m generate_word_report_v2.main
    或
    cd scripts/generate_word_report_v2
    python main.py
"""

import sys
import os
import json
import time
from datetime import datetime, timedelta
from docx import Document

# 添加项目路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))
sys.path.insert(0, PROJECT_DIR)

# 切换到项目根目录（确保数据文件路径正确）
os.chdir(PROJECT_DIR)

# 导入项目根目录的utils模块
from utils.config_loader import load_placement_config
from utils.analysis_tools import PrivatePlacementRiskAnalyzer
from utils.font_manager import get_font_prop

# 导入本模块的工具函数和章节模块
# 添加当前目录到路径以支持模块导入
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

import module_utils as utils
import chapter01_overview
import chapter02_valuation
import chapter03_dcf
import chapter04_sensitivity
import chapter05_montecarlo
import chapter06_scenario
import chapter07_stress
import chapter08_var
import chapter09_01_evaluation
import chapter09_advice
import appendix_scenarios

# 配置路径
# 数据文件在当前项目的data/目录下（PROJECT_DIR已经指向price_maintenance_risk_analysis）
DATA_DIR = os.path.join(PROJECT_DIR, 'data')
REPORTS_DIR = os.path.join(PROJECT_DIR, 'reports')
IMAGES_DIR = os.path.join(REPORTS_DIR, 'images')
OUTPUTS_DIR = os.path.join(REPORTS_DIR, 'outputs')

# 确保目录存在
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(OUTPUTS_DIR, exist_ok=True)

# 获取中文字体
font_prop = get_font_prop()


def _get_historical_price_and_ma20(stock_code, issue_date_str, force_regenerate=False):
    """
    获取指定日期的股票价格和MA20（基于发行日）

    重要：一旦发行日确定，MA20等基准数据将被锁定保存到DB，后续只更新当前价格

    参数:
        stock_code: 股票代码
        issue_date_str: 发行日/报价日字符串（YYYYMMDD格式）
        force_regenerate: 是否强制重新生成锁定数据（默认False）

    返回:
        tuple: (bidding_date_price, ma20_price) 或 (None, None)
    """
    try:
        import sys
        import os
        from datetime import datetime
        scripts_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'scripts')
        if scripts_dir not in sys.path:
            sys.path.insert(0, scripts_dir)

        try:
            from utils.db_manager import ValuationDB
            db = ValuationDB()
        except ImportError:
            db = None

        # 检查DB中是否存在锁定的发行日数据
        if db and not force_regenerate:
            locked_data = db.load_issue_date_locked(stock_code, issue_date_str)
            if locked_data:
                print(f"  ✅ 使用已锁定的发行日数据(DB)（发行日：{issue_date_str}）")
                ma20_price = locked_data.get('ma_20')
                bidding_date_price = locked_data.get('issue_date_price')
                # 回溯分析模式：current_price 统一为报价日价格，不更新到最新价
                # （所有指标都基于报价日，current_price也应是报价日价格）
                return bidding_date_price, ma20_price

        # DB中无数据或需要重新生成
        from update_market_data import generate_market_data, fetch_latest_data

        print(f"  基于发行日{issue_date_str}生成市场数据...")
        market_data = generate_market_data(stock_code, stock_code, issue_date_str)

        if market_data is None:
            return None, None

        ma20_price = market_data.get('ma_20')
        bidding_date_price = market_data.get('current_price')

        locked_data = {
            'issue_date': issue_date_str,
            'issue_date_price': bidding_date_price,
            'ma_20': ma20_price,
            'current_price': bidding_date_price,
            'analysis_date': issue_date_str,
            'stock_code': stock_code,
            'locked_timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

        # 保存到DB
        if db:
            db.save_issue_date_locked(stock_code, issue_date_str, locked_data)
            print(f"  ✅ 发行日数据已锁定保存(DB)")
        else:
            # 回退到文件
            locked_data_file = os.path.join(DATA_DIR, f"{stock_code.replace('.', '_')}_issue_date_locked.json")
            with open(locked_data_file, 'w', encoding='utf-8') as f:
                json.dump(locked_data, f, ensure_ascii=False, indent=2)
            print(f"  ✅ 发行日数据已锁定保存(文件)")

        print(f"     发行日：{issue_date_str}")
        print(f"     MA20价格：{ma20_price:.2f}元")
        print(f"     发行日价格：{bidding_date_price:.2f}元")

        # 自动生成基于发行日的锁定指数数据
        print()
        print(f"  🔄 生成基于发行日的锁定指数数据...")
        try:
            from update_indices_data import rebuild_locked_indices_data
            indices_results = rebuild_locked_indices_data(stock_code, issue_date_str)
            if indices_results:
                print(f"  ✅ 锁定指数数据生成成功")
            else:
                print(f"  ⚠️ 锁定指数数据生成失败，将使用最新指数数据")
        except Exception as e:
            print(f"  ❌ 锁定指数数据生成异常: {e}")
            print(f"  ⚠️ 将使用最新指数数据")

        return bidding_date_price, ma20_price

    except Exception as e:
        print(f"  ❌ 获取历史价格和MA20失败: {e}")
        import traceback
        traceback.print_exc()
        return None, None


def check_data_freshness(market_data, max_trading_days=0):
    """
    检查数据是否过期（基于交易日历）

    参数:
        market_data: 市场数据字典，应包含'analysis_date'
        max_trading_days: 允许的最大落后交易日数（默认0天，不允许落后）

    返回:
        bool: 数据是否新鲜（未过期）
        str: 信息/警告信息
    """
    if 'analysis_date' not in market_data:
        return True, "数据文件中未找到analysis_date字段"

    try:
        analysis_date_str = market_data['analysis_date']
        analysis_date = datetime.strptime(analysis_date_str, '%Y%m%d')
        current_date = datetime.now()

        # 计算交易日差距（排除周末）
        trading_days_diff = 0
        temp_date = analysis_date
        while temp_date.date() < current_date.date():  # 只比较日期部分，忽略时间
            temp_date += timedelta(days=1)
            # 只计算周一到周五（0-4），排除周六周日（5-6）
            if temp_date.weekday() < 5:  # 0=周一, 4=周五
                trading_days_diff += 1

        # 检查宽容期：如果交易日差距为1天且当前时间在20:00之前，给予宽容
        if trading_days_diff == 1 and current_date.hour < 20:
            return True, f"数据新鲜（当日20:00前，使用{analysis_date_str}数据）"

        # 如果当天是周末或节假日，可能没有最新数据，给予宽容
        current_weekday = current_date.weekday()
        if current_weekday >= 5:  # 周末
            # 周末时多给2天宽容期
            trading_days_diff = max(0, trading_days_diff - 2)

        if trading_days_diff > max_trading_days:
            warning_msg = f"""
数据过期警告：
   数据文件日期：{analysis_date_str}（{analysis_date.strftime('%Y年%m月%d日')} {['周一','周二','周三','周四','周五','周六','周日'][analysis_date.weekday()]}）
   当前系统日期：{current_date.strftime('%Y%m%d')}（{current_date.strftime('%Y年%m月%d日')} {['周一','周二','周三','周四','周五','周六','周日'][current_weekday]}）
   交易日落后：{trading_days_diff}个交易日
   最大允许：{max_trading_days}个交易日

   建议：请运行以下命令更新数据：
   python scripts/update_indices_data.py

   注意：股票价格实时变动，使用过期数据可能导致分析结果不准确！
            """
            return False, warning_msg.strip()
        else:
            if trading_days_diff == 0:
                return True, "数据最新，无交易日落后"
            else:
                return True, f"数据新鲜，落后{trading_days_diff}个交易日（在{max_trading_days}个交易日允许范围内）"

    except Exception as e:
        return True, f"无法解析数据日期：{e}"


def generate_report(stock_code='300735.SZ', stock_name='光弘科技', issue_date=None, force=False, debug_mode=False, target_chapter=None):
    """
    生成完整的定增风险分析报告

    参数:
        stock_code: 股票代码
        stock_name: 股票名称
        issue_date: 报价日（字符串格式，如'20240315'），可选
        force: 是否强制使用过期数据（跳过用户确认），默认False

    返回:
        document: Word文档对象
    """
    import sys
    from datetime import datetime, timedelta
    print("="*70)
    print(" 定增风险分析报告生成器（模块化版本 V2.0）")
    print("="*70)
    print(f"股票代码: {stock_code}")
    print(f"股票名称: {stock_name}")
    print(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    if debug_mode:
        print("🔍 调试模式已启用")
    if target_chapter:
        print(f"📄 仅生成章节: {target_chapter}")

    print("="*70)

    # 创建Word文档
    document = Document()
    utils.setup_chinese_font(document)

    # 检查市场数据是否存在（优先DB，其次文件）
    from utils.db_manager import ValuationDB
    db = ValuationDB()
    db_market = db.load_market_data(stock_code)
    market_data_file = os.path.join(DATA_DIR, f"{stock_code.replace('.', '_')}_market_data.json")

    if not db_market and not os.path.exists(market_data_file):
        print(f"⚠️ DB和文件中均无市场数据，自动生成...")

        try:
            # 导入update_market_data模块
            scripts_dir = os.path.join(PROJECT_DIR, 'scripts')
            if scripts_dir not in sys.path:
                sys.path.insert(0, scripts_dir)

            # 使用模块导入方式，避免命名冲突
            import importlib
            update_module = importlib.import_module('update_market_data')

            # 生成市场数据
            print(f"   调用update_market_data.generate_market_data...")
            updated_market_data = update_module.generate_market_data(stock_code, stock_name, issue_date)

            if updated_market_data:
                # 优先保存到DB
                print(f"   保存数据到DB...")
                db.save_market_data(stock_code, updated_market_data)
                print(f"✅ 市场数据生成成功！已保存到DB")
                print(f"   数据日期: {updated_market_data.get('analysis_date', 'N/A')}")
                print(f"   当前价格: {updated_market_data.get('current_price', 'N/A')}")
                print(f"   MA20: {updated_market_data.get('ma_20', 'N/A')}")
            else:
                print(f"❌ 市场数据生成失败！")
                print(f"   请检查网络连接和Tushare API配置")
                sys.exit(1)

        except Exception as e:
            import traceback
            print(f"❌ 自动生成市场数据失败: {e}")
            print(f"   详细错误信息:")
            traceback.print_exc()
            print(f"   请手动运行以下命令生成数据：")
            print(f"   python scripts/update_market_data.py --stock {stock_code} --name {stock_name}")
            sys.exit(1)

    # 加载配置数据
    print("\n 加载配置数据...")
    project_params, risk_params, market_data = load_placement_config(stock_code)
    industry_data = _load_industry_data(stock_code, issue_date=issue_date)

    # 优先从配置文件中读取发行日（一旦配置文件中有发行日，就始终使用它）
    config_issue_date = project_params.get('issue_date', '')
    if config_issue_date:
        if issue_date and config_issue_date != issue_date:
            print(f"  ⚠️ 配置文件发行日({config_issue_date})与命令行参数({issue_date})不同，使用配置文件中的发行日")
        else:
            print(f"  从配置文件读取发行日：{config_issue_date}")
        issue_date = config_issue_date  # 始终使用配置文件中的发行日

    # 处理报价日和询价邀请日
    if issue_date:
        # 情况1：指定报价日（需要获取历史价格）
        from datetime import datetime, timedelta
        issue_date_obj = datetime.strptime(issue_date, '%Y%m%d')
        invitation_date = issue_date_obj - timedelta(days=3)
        invitation_date_str = invitation_date.strftime('%Y%m%d')

        # 获取报价日当天的股票价格和MA20（基于发行日）
        print(f" 获取报价日{issue_date}的历史价格和MA20（基于发行日{issue_date}）...")
        bidding_date_price, ma20_price = _get_historical_price_and_ma20(stock_code, issue_date)

        if bidding_date_price is None or ma20_price is None:
            print(f"错误：无法获取报价日{issue_date}的历史价格或MA20")
            print(f"请检查网络连接和数据可用性")
            sys.exit(1)

        # 保存原始的current_price（最新交易日价格），不要覆盖
        original_current_price = market_data['current_price']

        # 更新market_data中的MA20，但保留current_price为最新交易日价格
        market_data['ma_20'] = ma20_price
        # market_data['current_price'] 保持为最新交易日价格，不覆盖

        project_params['ma20_from_issue'] = ma20_price
        project_params['current_price'] = bidding_date_price  # 用于全文实际溢价率计算
        project_params['bidding_date_price'] = bidding_date_price
        project_params['latest_trading_date_price'] = original_current_price  # 保存最新交易日价格

        print(f"  MA20价格（基于发行日{issue_date}的成交量加权均价）：{ma20_price:.2f}元")
        print(f"  发行日价格（用于实际溢价率计算）：{bidding_date_price:.2f}元")
        print(f"  最新交易日价格（保留在market_data中）：{original_current_price:.2f}元")

        project_params['issue_date'] = issue_date
        project_params['invitation_date'] = invitation_date_str
        project_params['invitation_date_fixed'] = True  # 标记为固定日期
        print(f" 使用指定报价日：{issue_date}")
        print(f" 询价邀请日：{project_params['invitation_date']}（报价日-3天，固定）")
        print(f" 报价日当天价格：{bidding_date_price:.2f}元")
    else:
        # 情况2：使用当前日期作为报价日
        from datetime import datetime, timedelta
        today = datetime.now()
        invitation_date = today - timedelta(days=3)
        invitation_date_str = invitation_date.strftime('%Y%m%d')

        # 如果没有具体报价日，使用当前日
        issue_date = today.strftime('%Y%m%d')

        # 获取当前价格和MA20（基于发行日）
        print(f" 计算当前价格和MA20（基于发行日{issue_date}）...")
        bidding_date_price, ma20_price = _get_historical_price_and_ma20(stock_code, issue_date)

        if bidding_date_price is None or ma20_price is None:
            print(f"错误：无法获取当前价格或MA20")
            print(f"请检查网络连接和数据可用性")
            sys.exit(1)

        # 保存原始的current_price（最新交易日价格），不要覆盖
        original_current_price = market_data['current_price']

        # 更新market_data中的MA20，但保留current_price为最新交易日价格
        market_data['ma_20'] = ma20_price
        # market_data['current_price'] 保持为最新交易日价格，不覆盖

        project_params['issue_date'] = issue_date
        project_params['invitation_date'] = invitation_date_str
        project_params['invitation_date_fixed'] = False  # 标记为动态日期
        project_params['bidding_date_price'] = bidding_date_price
        project_params['current_price'] = bidding_date_price  # 用于全文实际溢价率计算
        project_params['ma20_from_issue'] = ma20_price
        project_params['latest_trading_date_price'] = original_current_price  # 保存最新交易日价格

        print(f" 使用当前日期作为报价日：{issue_date}")
        print(f" 询价邀请日：{invitation_date_str}（报价日-3天，动态）")
        print(f" 报价日当天价格：{bidding_date_price:.2f}元")
        print(f"  MA20价格（基于发行日{issue_date}的成交量加权均价）：{ma20_price:.2f}元")
        print(f"  最新交易日价格（保留在market_data中）：{original_current_price:.2f}元")

    # 检查数据新鲜度
    print("\n 检查数据新鲜度...")
    is_fresh, data_msg = check_data_freshness(market_data)
    print(f" {data_msg}")

    if not is_fresh:
        print("="*70)
        print(" 警告：数据已过期，自动更新数据...")
        print("="*70)

        # 自动更新数据（force模式或交互模式均自动更新）
        try:
            import subprocess

            # 更新市场指数数据
            update_indices_script = os.path.join(PROJECT_DIR, 'scripts', 'update_indices_data.py')

            if os.path.exists(update_indices_script):
                print(" 1. 更新市场指数数据...")
                result = subprocess.run([sys.executable, update_indices_script],
                                      capture_output=True,
                                      text=True,
                                      timeout=300)

                if result.returncode == 0:
                    print(" 市场指数数据更新成功！")
                else:
                    print(f" 市场指数数据更新失败: {result.stderr[:200]}")

            # 更新个股市场数据
            print(" 2. 更新个股市场数据...")
            scripts_dir = os.path.join(PROJECT_DIR, 'scripts')
            if scripts_dir not in sys.path:
                sys.path.insert(0, scripts_dir)
            from update_market_data import generate_market_data

            updated_market_data = generate_market_data(stock_code, stock_name, issue_date)

            if updated_market_data:
                # 保存到DB
                try:
                    from utils.db_manager import ValuationDB
                    db = ValuationDB()
                    db.save_market_data(stock_code, updated_market_data)
                except Exception:
                    pass

                print(f" 个股市场数据更新成功！")

                # 重新加载数据
                print("\n 重新加载更新后的数据...")
                project_params, risk_params, market_data = load_placement_config(stock_code)
                industry_data = _load_industry_data(stock_code, issue_date=issue_date)

                is_fresh_updated, data_msg_updated = check_data_freshness(market_data)
                print(f" {data_msg_updated}")
            else:
                print(" 个股市场数据更新失败！使用现有数据继续...")
        except subprocess.TimeoutExpired:
            print(" 数据更新超时，使用现有数据继续...")
        except Exception as e:
            print(f" 数据更新出错: {e}，使用现有数据继续...")

        print("="*70)

    # 创建分析器（注意参数顺序：issue_price, issue_shares, lockup_period, current_price, risk_free_rate）
    analyzer = PrivatePlacementRiskAnalyzer(
        project_params['issue_price'],
        project_params['issue_shares'],
        project_params['lockup_period'],
        project_params['current_price'],
        project_params.get('risk_free_rate', 0.03)
    )

    # 共享数据字典（在各章节间传递）
    context = {
        'stock_code': stock_code,
        'stock_name': stock_name,
        'project_params': project_params,
        'risk_params': risk_params,  # 添加risk_params
        'market_data': market_data,
        'industry_data': industry_data,
        'analyzer': analyzer,
        'document': document,
        'font_prop': font_prop,
        'IMAGES_DIR': IMAGES_DIR,
        'DATA_DIR': DATA_DIR,
        'OUTPUTS_DIR': OUTPUTS_DIR,
        # 各章节计算结果（供后续章节使用）
        'results': {}
    }

    # ==================== 依次调用各章节 ====================
    # 第一章：项目概况（包含封面和目录）
    print("\n 生成第一章：项目概况...")
    context = chapter01_overview.generate_chapter(context)

    # 第一章完成后，设置页眉（从正文开始有页眉）
    company_name = project_params.get('company_name', stock_name)

    # 只在第一章后设置一次页眉，不创建多个section
    # 使用增强的统一页眉显示公司名称+股票代码+报告标题+日期
    utils.setup_document_header(document, company_name, stock_code)

    # 可选：使用方案5（奇偶页不同页眉）
    # 奇数页显示报告标题，偶数页显示章节标题
    # 如果需要使用方案5，请取消下面一行的注释，并注释掉上面的统一页眉设置
    # utils.setup_odd_even_headers(document, company_name, stock_code)

    # 第二章：相对估值分析
    print("\n 生成第二章：相对估值分析...")
    context = chapter02_valuation.generate_chapter(context)
    # 如果使用方案5奇偶页眉，可在此处更新偶数页页眉：
    # utils.update_even_page_header(document, "第二章 相对估值分析")

    # 第三章：DCF估值分析
    print("\n 生成第三章：DCF估值分析...")
    context = chapter03_dcf.generate_chapter(context)
    # 如果使用方案5奇偶页眉，可在此处更新偶数页页眉：
    # utils.update_even_page_header(document, "第三章 DCF估值分析")

    # 第四章：敏感性分析
    print("\n 生成第四章：敏感性分析...")
    context = chapter04_sensitivity.generate_chapter(context)

    # 第五章：蒙特卡洛模拟
    print("\n 生成第五章：蒙特卡洛模拟...")
    context = chapter05_montecarlo.generate_chapter(context)

    # 第六章：情景分析
    print("\n 生成第六章：情景分析...")
    context = chapter06_scenario.generate_chapter(context)

    # 第七章：压力测试
    print("\n 生成第七章：压力测试...")
    context = chapter07_stress.generate_chapter(context)

    # 第八章：VaR风险度量
    print("\n 生成第八章：VaR风险度量...")
    context = chapter08_var.generate_chapter(context)

    # 第九章：风控建议与风险提示
    print("\n 生成第九章：风控建议与风险提示...")
    # 生成第九章全部内容（包括9.1综合评估汇总和9.2-9.6其他章节）
    context = chapter09_advice.generate_chapter(context)
    # 如果使用方案5奇偶页眉，可在此处更新偶数页页眉：
    # utils.update_even_page_header(document, "第九章 风控建议与风险提示")

    # 附件：情景数据表
    print("\n 生成附件：情景数据表...")
    context = appendix_scenarios.generate_chapter(context)

    # 添加页码
    print("\n 添加页码...")
    utils.add_page_numbers(document)

    print("\n" + "="*70)
    print(" 报告生成完成！")
    print("="*70)

    return document


def _ensure_market_data(stock_code, stock_name, market_data_file, issue_date=None):
    """确保市场数据存在且包含 ma_20，缺失或无效时自动重新生成。

    陈旧检测：通过 schema_version 字段判断数据版本，缺失或版本过旧则重建。
    """
    # 先尝试从DB加载
    try:
        from utils.db_manager import ValuationDB
        db = ValuationDB()
        existing = db.load_market_data(stock_code)
        if existing and existing.get('ma_20'):
            # 陈旧检测：缺 schema_version 或版本不是最新的，重新生成
            if existing.get('schema_version') != 'v2_issue_date':
                print(f"  ⚠️ 检测到旧版市场数据(版本={existing.get('schema_version', '无')}),重新生成...")
            else:
                print(f"  市场数据(DB)已存在，无需重新生成")
                return
    except Exception:
        pass

    # DB中无有效数据或数据陈旧，检查文件
    use_file_data = False
    if os.path.exists(market_data_file):
        try:
            with open(market_data_file, 'r', encoding='utf-8') as f:
                existing = json.load(f)
            if existing and existing.get('ma_20') and existing.get('schema_version') == 'v2_issue_date':
                # 文件有新数据但DB没有，导入到DB
                try:
                    from utils.db_manager import ValuationDB
                    db = ValuationDB()
                    db.save_market_data(stock_code, existing)
                    print(f"  市场数据已从文件导入DB")
                except Exception:
                    pass
                return
        except Exception:
            pass

    # 需要重新生成
    scripts_dir = os.path.join(PROJECT_DIR, 'scripts')
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import importlib
    update_module = importlib.import_module('update_market_data')
    updated_data = update_module.generate_market_data(
        stock_code, stock_name or stock_code, issue_date,
    )
    if updated_data:
        # 保存到DB
        try:
            from utils.db_manager import ValuationDB
            db = ValuationDB()
            db.save_market_data(stock_code, updated_data)
        except Exception:
            pass
        print(f"  市场数据已自动生成并保存到DB")
    else:
        raise ValueError(
            f"无法生成 {stock_code} 的市场数据，请检查Tushare API配置"
        )


def generate_report_headless(stock_code, stock_name=None, issue_date=None, force=True):
    """无头模式：运行完整分析管线，返回结构化决策结果，不保存Word文档。

    Args:
        stock_code: 股票代码 (如 '300735.SZ')
        stock_name: 股票名称（可选，用于显示）
        issue_date: 报价日 (YYYYMMDD格式，可选)
        force: 强制使用过期数据（默认True，批量模式下跳过确认）

    Returns:
        dict with keys:
            stock_code, stock_name,
            decision_conclusion: dict from calculate_decision_conclusion(),
            error: str or None
    """
    result = {
        'stock_code': stock_code,
        'stock_name': stock_name or stock_code,
        'decision_conclusion': None,
        'error': None,
    }

    try:
        print(f"\n[headless] 分析 {stock_code} ({stock_name or ''})...")

        # 确保 market_data 文件存在且有效（必须在 load_placement_config 之前）
        import io as _io
        import contextlib as _ctx
        import sys as _sys
        import threading as _th

        # ── 线程安全的 stdout 代理(修多线程下 redirect_stdout 互相覆盖致输出冻结) ──
        # 安装一次: sys.stdout → _Proxy, 按线程局部 sink 决定写哪(默认真实终端)。
        # 各线程用 _suppress() 把"自己"的输出暂存 StringIO(静默章节细节), 不影响其它线程。
        if not getattr(generate_report_headless, '_proxy_on', False):
            _real_out = _sys.stdout
            _sink_tls = _th.local()

            class _Proxy:
                def write(self, data):
                    s = getattr(_sink_tls, 'sink', None)
                    (s if s is not None else _real_out).write(data)
                def flush(self):
                    s = getattr(_sink_tls, 'sink', None)
                    (s if s is not None else _real_out).flush()

            _sys.stdout = _Proxy()
            generate_report_headless._proxy_on = True
            generate_report_headless._sink_tls = _sink_tls
            generate_report_headless._real_out = _real_out

        @_ctx.contextmanager
        def _suppress():
            """线程局部静默: 把当前线程的 print 暂存 StringIO, 不影响其它线程/主线程。"""
            _tls = generate_report_headless._sink_tls
            prev = getattr(_tls, 'sink', None)
            _tls.sink = _io.StringIO()
            try:
                yield
            finally:
                _tls.sink = prev

        market_data_file = os.path.join(DATA_DIR, f"{stock_code.replace('.', '_')}_market_data.json")
        # 市场数据生成：抑制详细输出
        print(f"  📊 生成市场数据...", flush=True)
        with _suppress():
            _ensure_market_data(stock_code, stock_name, market_data_file, issue_date)

        # 加载配置
        project_params, risk_params, market_data = load_placement_config(stock_code)
        # 关键：把issue_date存入project_params，供各章节(Ch3等)回溯分析使用
        if issue_date:
            project_params['issue_date'] = issue_date
        if not stock_name:
            stock_name = project_params.get('company_name', stock_code)
            result['stock_name'] = stock_name

        # 加载行业数据（失败时用空dict，不阻塞后续分析）— 抑制输出
        with _suppress():
            industry_data = _load_industry_data(stock_code, issue_date=issue_date)
        if industry_data is None:
            print(f"  ⚠️ 行业数据加载失败，将跳过行业相关分析")
            industry_data = {}

        # 锁定发行日价格
        if issue_date:
            _get_historical_price_and_ma20(stock_code, issue_date)

        # 创建分析器
        analyzer = PrivatePlacementRiskAnalyzer(
            project_params['issue_price'],
            project_params['issue_shares'],
            project_params['lockup_period'],
            project_params['current_price'],
            project_params.get('risk_free_rate', 0.03),
        )

        # 创建空Document作为占位（不保存）
        document = Document()

        context = {
            'stock_code': stock_code,
            'stock_name': stock_name,
            'project_params': project_params,
            'risk_params': risk_params,
            'market_data': market_data,
            'industry_data': industry_data,
            'analyzer': analyzer,
            'document': document,
            'font_prop': font_prop,
            'IMAGES_DIR': IMAGES_DIR,
            'DATA_DIR': DATA_DIR,
            'OUTPUTS_DIR': OUTPUTS_DIR,
            'results': {},
        }

        # headless模式：跳过图表生成 + 抑制内部输出（批量筛选只需计时和决策）
        import matplotlib
        matplotlib.use('Agg')  # 确保非交互后端
        import matplotlib.pyplot as _plt
        _plt.savefig = lambda *a, **kw: None  # 禁用savefig，跳过所有图表渲染
        import io as _io
        import contextlib as _ctx

        print(f"  🔍 分析中...", flush=True)

        # 依次调用各章节（抑制详细输出，只保留计时）
        import io as _io
        import contextlib as _ctx
        chapters = [
            ('Ch1 概况', chapter01_overview.generate_chapter),
            ('Ch2 相对估值', chapter02_valuation.generate_chapter),
            ('Ch3 DCF估值', chapter03_dcf.generate_chapter),
            ('Ch4 敏感性', chapter04_sensitivity.generate_chapter),
            ('Ch5 蒙特卡洛', chapter05_montecarlo.generate_chapter),
            ('Ch6 情景分析', chapter06_scenario.generate_chapter),
            ('Ch7 压力测试', chapter07_stress.generate_chapter),
            ('Ch8 VaR', chapter08_var.generate_chapter),
            ('Ch9.1 评估汇总', chapter09_01_evaluation.generate_chapter),
            ('Ch9 风控建议', chapter09_advice.generate_chapter),
        ]
        for ch_name, ch_fn in chapters:
            try:
                t_ch = time.time()
                with _suppress():
                    context = ch_fn(context)
                ch_elapsed = time.time() - t_ch
                print(f"  ⏱ {ch_name} {ch_elapsed:.1f}s", flush=True)
            except Exception as ch_err:
                print(f"  ⚠️ {ch_name} 失败: {ch_err}", flush=True)

        # 提取决策结论
        decision = context.get('results', {}).get('decision_conclusion')
        result['decision_conclusion'] = decision

        # 决策先打印（用户看到进度，不被DB写入卡住）
        if decision:
            print(f"  → 决策: {decision['decision']}", flush=True)
        else:
            print(f"  → 决策结论未生成", flush=True)

        # 章节执行后保存更新过的数据到DB（FCF/财报等在Ch3才生成）
        # DB写入放在决策输出之后，避免用户以为卡住
        try:
            from utils.db_manager import ValuationDB
            db = ValuationDB()
            hfcf = project_params.get('historical_fcf_data', {})
            if hfcf and hfcf.get('data'):
                db.save_historical_fcf(stock_code, hfcf['data'])
            db.save_placement_params(stock_code, project_params)
        except Exception:
            pass

    except Exception as e:
        result['error'] = str(e)
        print(f"  ⚠️ 分析失败: {e}", flush=True)

    return result


def _load_industry_data(stock_code, auto_generate=True, issue_date=None):
    """
    加载行业数据，支持自动生成和数据更新检查

    参数:
        stock_code: 股票代码
        auto_generate: 是否自动生成缺失的数据（默认True）

    返回:
        行业数据字典，如果失败返回None
    """
    import json
    from datetime import datetime, timedelta

    try:
        from utils.db_manager import ValuationDB
        db = ValuationDB()
    except ImportError:
        db = None

    # 优先从DB加载
    if db:
        industry_data = db.load_industry_data(stock_code)
        if industry_data:
            # 检查数据新鲜度
            generated_at = industry_data.get('generated_at', '')
            if generated_at:
                try:
                    data_time = datetime.strptime(generated_at[:10], '%Y-%m-%d')
                    data_age = (datetime.now() - data_time).days
                    if data_age > 7:
                        print(f" ⚠️ 行业数据已过期 {data_age} 天（生成时间: {generated_at}）")
                        if auto_generate:
                            print(f" 📥 自动更新行业数据...")
                            try:
                                scripts_dir = os.path.join(PROJECT_DIR, 'scripts')
                                if scripts_dir not in sys.path:
                                    sys.path.insert(0, scripts_dir)
                                from update_market_data import generate_industry_data
                                new_data = generate_industry_data(stock_code, issue_date=issue_date)
                                if new_data:
                                    db.save_industry_data(stock_code, new_data)
                                    print(f" ✅ 行业数据更新成功！")
                                    return new_data
                                else:
                                    print(f" ⚠️ 行业数据更新失败，使用旧数据")
                            except Exception as e:
                                print(f" ❌ 自动更新行业数据失败: {e}，使用旧数据")
                    else:
                        print(f" ✅ 已加载行业数据(DB): {stock_code}")
                except (ValueError, TypeError):
                    print(f" ✅ 已加载行业数据(DB): {stock_code}")
            else:
                print(f" ✅ 已加载行业数据(DB): {stock_code}")
            return industry_data

    # DB中无数据，尝试文件
    industry_data_file = os.path.join(DATA_DIR, f"{stock_code.replace('.', '_')}_industry_data.json")

    if os.path.exists(industry_data_file):
        with open(industry_data_file, 'r', encoding='utf-8') as f:
            industry_data = json.load(f)

        # 导入到DB
        if db and industry_data:
            try:
                db.save_industry_data(stock_code, industry_data)
                print(f" ✅ 行业数据已从文件导入DB")
            except Exception:
                pass

        print(f" ✅ 已加载行业数据(文件): {industry_data_file}")
        return industry_data

    # 既无DB数据也无文件，自动生成
    if auto_generate:
        print(f" ⚠️ 未找到行业数据，自动生成...")
        try:
            scripts_dir = os.path.join(PROJECT_DIR, 'scripts')
            if scripts_dir not in sys.path:
                sys.path.insert(0, scripts_dir)
            from update_market_data import generate_industry_data
            industry_data = generate_industry_data(stock_code, issue_date=issue_date)
            if industry_data:
                # 保存到DB
                if db:
                    db.save_industry_data(stock_code, industry_data)
                print(f" ✅ 行业数据生成成功！")
                return industry_data
            else:
                print(f" ❌ 行业数据生成失败！")
                return None
        except Exception as e:
            print(f" ❌ 自动生成行业数据失败: {e}")
            return None
    else:
        return None


def save_report(document, stock_code, stock_name):
    """保存报告到文件"""
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{stock_name}_{stock_code}_定增市场风险分析报告_{timestamp}.docx"
    output_path = os.path.join(OUTPUTS_DIR, filename)

    document.save(output_path)
    print(f"\n 报告已保存到: {output_path}")
    return output_path


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='定增风险分析报告生成器（模块化版本）')
    parser.add_argument('--stock', type=str, default='300735.SZ', help='股票代码（默认：300735.SZ）')
    parser.add_argument('--name', type=str, default='光弘科技', help='股票名称（默认：光弘科技）')
    parser.add_argument('--issue-date', type=str, default=None, help='报价日（格式：YYYYMMDD，如20240315，可选）')
    parser.add_argument('--output', type=str, default=None, help='输出文件名（可选）')
    parser.add_argument('--force', action='store_true', help='强制使用过期数据，跳过用户确认')

    args = parser.parse_args()

    # 验证报价日格式
    if args.issue_date:
        try:
            from datetime import datetime
            # 验证日期格式为YYYYMMDD
            datetime.strptime(args.issue_date, '%Y%m%d')
        except ValueError:
            print(f"❌ 错误：报价日格式不正确 '{args.issue_date}'")
            print("请使用格式：YYYYMMDD（例如：20240315）")
            sys.exit(1)

    # 生成报告
    doc = generate_report(args.stock, args.name, args.issue_date, args.force)

    # 保存报告
    if args.output:
        output_path = os.path.join(OUTPUTS_DIR, args.output)
        doc.save(output_path)
        print(f"\n 报告已保存到: {output_path}")
    else:
        save_report(doc, args.stock, args.name)
