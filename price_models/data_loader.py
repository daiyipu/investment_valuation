"""从DB加载股价数据供模型使用"""

import sys
import os
import numpy as np
import pandas as pd

# 添加项目路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
PRICE_MAINT_DIR = os.path.join(PROJECT_ROOT, 'price_maintenance_risk_analysis')
sys.path.insert(0, PRICE_MAINT_DIR)


def load_price_series(stock_code, db_path=None, window=None):
    """从DB加载股价序列

    Args:
        stock_code: 股票代码，如 '300735.SZ'
        db_path: 数据库路径（None=默认）
        window: 截取最近N个交易日（None=全部）

    Returns:
        pd.Series 或 None
    """
    try:
        from utils.db_manager import ValuationDB
        db = ValuationDB(db_path=db_path)
        md = db.load_market_data(stock_code)
        if not md:
            print(f"⚠️ DB中无 {stock_code} 的市场数据")
            return None

        prices = md.get('price_series', [])
        if not prices or len(prices) < 100:
            print(f"⚠️ {stock_code} 价格数据不足: {len(prices) if prices else 0} 条")
            return None

        series = pd.Series(prices)
        if window and len(series) > window:
            series = series.iloc[-window:]

        return series

    except ImportError:
        print("⚠️ 无法导入 db_manager，请确认 price_maintenance_risk_analysis 在正确路径")
        return None


def load_log_returns(stock_code, db_path=None, window=None):
    """从DB加载对数收益率序列

    Returns:
        pd.Series 或 None
    """
    prices = load_price_series(stock_code, db_path=db_path, window=window)
    if prices is None:
        return None
    log_returns = np.log(prices).diff().dropna()
    return log_returns


def list_stocks_with_data(db_path=None):
    """列出DB中有市场数据的所有股票

    Returns:
        list of (stock_code, stock_name)
    """
    try:
        from utils.db_manager import ValuationDB
        db = ValuationDB(db_path=db_path)
        conn = db.get_connection()
        rows = conn.execute('''
            SELECT m.stock_code, s.stock_name
            FROM market_data m
            LEFT JOIN stocks s ON m.stock_code = s.stock_code
            ORDER BY m.stock_code
        ''').fetchall()
        return rows
    except ImportError:
        return []
