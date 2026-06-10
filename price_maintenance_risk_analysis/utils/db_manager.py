#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SQLite 数据库管理器

统一管理所有定增分析数据，替代原有的散乱 JSON 文件存储。
"""

import json
import os
import pymysql
from pymysql.cursors import DictCursor
from datetime import datetime, timedelta

DB_FILENAME = 'valuation.db'


class _MySQLConnAdapter:
    """适配层：让pymysql连接支持conn.execute()（兼容原sqlite3代码）。"""
    def __init__(self, conn):
        self._conn = conn
    def execute(self, sql, params=()):
        cur = self._conn.cursor()
        cur.execute(sql, params)
        return cur
    def executemany(self, sql, params):
        cur = self._conn.cursor()
        cur.executemany(sql, params)
        return cur
    def commit(self):
        self._conn.commit()
    def close(self):
        self._conn.close()
    @property
    def row_factory(self):
        return None
    @row_factory.setter
    def row_factory(self, val):
        pass  # DictCursor已经返回dict，忽略


class ValuationDB:
    """定增分析数据库访问层。"""

    # MySQL配置
    MYSQL_CONFIG = {
        'host': '127.0.0.1', 'port': 3306,
        'user': 'root', 'password': '',
        'database': 'investment_valuation', 'charset': 'utf8mb4',
    }

    def __init__(self, db_path=None):
        # db_path保留兼容(不再使用，改用MySQL)
        self.db_path = db_path or 'investment_valuation'

    # ==================== 基础操作 ====================

    def get_connection(self):
        conn = pymysql.connect(cursorclass=DictCursor, **self.MYSQL_CONFIG)
        return _MySQLConnAdapter(conn)

    def _ensure_db(self):
        """MySQL表已通过migrate_to_mysql.py创建，无需运行时建表。"""
        pass

    def _legacy_ensure_db(self):
        """旧SQLite建表（保留备用）。"""
        conn = self.get_connection()
        conn.executescript(SCHEMA)
        conn.commit()
        conn.close()

    # ==================== stocks ====================

    def upsert_stock(self, stock_code, stock_name=None, **kwargs):
        """插入或更新股票基本信息。"""
        conn = self.get_connection()
        existing = conn.execute(
            "SELECT stock_code FROM stocks WHERE stock_code=%s", (stock_code,)
        ).fetchone()

        fields = {'stock_name': stock_name}
        for k in ('sw_l1_code', 'sw_l1_name', 'sw_l2_code', 'sw_l2_name',
                   'sw_l3_code', 'sw_l3_name'):
            if k in kwargs:
                fields[k] = kwargs[k]

        if existing:
            sets = ', '.join(f"{k}=%s" for k in fields)
            vals = list(fields.values()) + [stock_code]
            conn.execute(f"UPDATE stocks SET {sets}, updated_at=NOW() WHERE stock_code=%s", vals)
        else:
            cols = ['stock_code'] + list(fields.keys())
            placeholders = ','.join(['%s'] * len(cols))
            vals = [stock_code] + list(fields.values())
            conn.execute(f"INSERT INTO stocks ({','.join(cols)}) VALUES ({placeholders})", vals)
        conn.commit()
        conn.close()

    def get_stock(self, stock_code):
        conn = self.get_connection()
        row = conn.execute("SELECT * FROM stocks WHERE stock_code=%s", (stock_code,)).fetchone()
        conn.close()
        return row if row else None

    # ==================== placement_params ====================

    def load_placement_params(self, stock_code):
        """加载定增参数配置，返回 dict（兼容原 JSON 格式，含 historical_fcf_data）。"""
        conn = self.get_connection()
        row = conn.execute(
            "SELECT * FROM placement_params WHERE stock_code=%s", (stock_code,)
        ).fetchone()
        if row is None:
            conn.close()
            return None

        result = row
        # 移除元数据字段
        result.pop('created_at', None)
        result.pop('updated_at', None)

        # 附加 historical_fcf_data
        fcf_rows = conn.execute(
            "SELECT * FROM historical_fcf WHERE stock_code=%s ORDER BY year", (stock_code,)
        ).fetchall()
        if fcf_rows:
            fcf_list = [r for r in fcf_rows]
            for item in fcf_list:
                item.pop('id', None)
                item.pop('stock_code', None)
            years = [r['year'] for r in fcf_list]
            result['historical_fcf_data'] = {
                'years': len(fcf_list),
                'year_range': [min(years), max(years)],
                'data': fcf_list,
            }
        else:
            result['historical_fcf_data'] = {'years': 5, 'year_range': [2020, 2024], 'data': []}

        # 添加 _notes
        result['_notes'] = {
            'financing_amount': '投资金额（元）- 固定1亿元',
            'lockup_period': '锁定期（月）- 默认6个月',
            'pricing_method': '定价方式',
        }

        conn.close()
        return result

    def save_placement_params(self, stock_code, params):
        """保存定增参数配置（upsert）。"""
        conn = self.get_connection()
        self.upsert_stock(stock_code, params.get('stock_name'))

        cols = [
            'stock_code', 'financing_amount', 'lockup_period', 'pricing_method',
            'premium_rate', 'risk_free_rate', 'net_assets', 'total_debt',
            'net_income', 'revenue_growth', 'operating_margin', 'beta',
        ]
        vals = [params.get(c.split('.')[-1]) for c in cols]
        vals[0] = stock_code  # 确保 stock_code 正确

        conn.execute(f"""
            INSERT INTO placement_params ({','.join(cols)})
            VALUES ({','.join(['%s'] * len(cols))})
            ON DUPLICATE KEY UPDATE
                financing_amount=VALUES(financing_amount),
                lockup_period=VALUES(lockup_period),
                pricing_method=VALUES(pricing_method),
                premium_rate=VALUES(premium_rate),
                risk_free_rate=VALUES(risk_free_rate),
                net_assets=VALUES(net_assets),
                total_debt=VALUES(total_debt),
                net_income=VALUES(net_income),
                revenue_growth=VALUES(revenue_growth),
                operating_margin=VALUES(operating_margin),
                beta=VALUES(beta),
                updated_at=NOW()
        """, vals)

        # 保存 historical_fcf_data（如果有）
        hfcf = params.get('historical_fcf_data', {})
        fcf_list = hfcf.get('data', [])
        if fcf_list:
            self.save_historical_fcf(stock_code, fcf_list, conn=conn)

        conn.commit()
        conn.close()

    # ==================== historical_fcf ====================

    def save_historical_fcf(self, stock_code, fcf_list, conn=None):
        """批量 upsert 历史FCF数据。"""
        close_after = False
        if conn is None:
            conn = self.get_connection()
            close_after = True

        for item in fcf_list:
            conn.execute("""
                INSERT INTO historical_fcf (stock_code, year, revenue, operate_profit,
                    net_income, nopat, depreciation, capex, wc_change, fcf)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    revenue=VALUES(revenue), operate_profit=VALUES(operate_profit),
                    net_income=VALUES(net_income), nopat=VALUES(nopat),
                    depreciation=VALUES(depreciation), capex=VALUES(capex),
                    wc_change=VALUES(wc_change), fcf=VALUES(fcf)
            """, (
                stock_code, item.get('year'), item.get('revenue'),
                item.get('operate_profit'), item.get('net_income'),
                item.get('nopat'), item.get('depreciation'),
                item.get('capex'), item.get('wc_change'), item.get('fcf'),
            ))

        if close_after:
            conn.commit()
            conn.close()

    # ==================== market_data ====================

    def load_market_data(self, stock_code):
        """加载市场数据，返回 dict（兼容原 JSON 格式）。"""
        conn = self.get_connection()
        row = conn.execute("SELECT * FROM market_data WHERE stock_code=%s", (stock_code,)).fetchone()
        conn.close()
        if row is None:
            return None
        result = row
        # 反序列化 JSON 字段
        for key in ('price_series', 'market_turnover'):
            if result.get(key) and isinstance(result[key], str):
                result[key] = json.loads(result[key])
        return result

    def save_market_data(self, stock_code, data):
        """保存市场数据（upsert）。"""
        conn = self.get_connection()
        self.upsert_stock(stock_code, data.get('stock_name'))

        # 序列化 JSON 字段
        price_series = data.get('price_series', [])
        market_turnover = data.get('market_turnover', {})
        if isinstance(price_series, list):
            price_series = json.dumps(price_series)
        if isinstance(market_turnover, dict):
            market_turnover = json.dumps(market_turnover)

        conn.execute("""
            INSERT INTO market_data (
                stock_code, analysis_date, latest_trading_date, issue_date, invitation_date,
                current_price, avg_price_all, median_price, price_std,
                volatility_20d, volatility_60d, volatility_120d, volatility_250d,
                annual_return_20d, annual_return_60d, annual_return_120d, annual_return_250d,
                period_return_20d, period_return_60d, period_return_120d, period_return_250d,
                ma_20, ma_30, ma_60, ma_120, ma_250,
                win_rate_20d, win_rate_60d, win_rate_120d, win_rate_250d,
                total_days, drift, volatility,
                price_series, market_turnover,
                data_source, generated_at
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON DUPLICATE KEY UPDATE
                analysis_date=VALUES(analysis_date),
                latest_trading_date=VALUES(latest_trading_date),
                issue_date=VALUES(issue_date),
                invitation_date=VALUES(invitation_date),
                current_price=VALUES(current_price),
                avg_price_all=VALUES(avg_price_all),
                median_price=VALUES(median_price),
                price_std=VALUES(price_std),
                volatility_20d=VALUES(volatility_20d),
                volatility_60d=VALUES(volatility_60d),
                volatility_120d=VALUES(volatility_120d),
                volatility_250d=VALUES(volatility_250d),
                annual_return_20d=VALUES(annual_return_20d),
                annual_return_60d=VALUES(annual_return_60d),
                annual_return_120d=VALUES(annual_return_120d),
                annual_return_250d=VALUES(annual_return_250d),
                period_return_20d=VALUES(period_return_20d),
                period_return_60d=VALUES(period_return_60d),
                period_return_120d=VALUES(period_return_120d),
                period_return_250d=VALUES(period_return_250d),
                ma_20=VALUES(ma_20), ma_30=VALUES(ma_30),
                ma_60=VALUES(ma_60), ma_120=VALUES(ma_120), ma_250=VALUES(ma_250),
                win_rate_20d=VALUES(win_rate_20d),
                win_rate_60d=VALUES(win_rate_60d),
                win_rate_120d=VALUES(win_rate_120d),
                win_rate_250d=VALUES(win_rate_250d),
                total_days=VALUES(total_days),
                drift=VALUES(drift),
                volatility=VALUES(volatility),
                price_series=VALUES(price_series),
                market_turnover=VALUES(market_turnover),
                data_source=VALUES(data_source),
                generated_at=VALUES(generated_at)
        """, (
            stock_code,
            data.get('analysis_date'), data.get('latest_trading_date'),
            data.get('issue_date'), data.get('invitation_date'),
            data.get('current_price'), data.get('avg_price_all'),
            data.get('median_price'), data.get('price_std'),
            data.get('volatility_20d'), data.get('volatility_60d'),
            data.get('volatility_120d'), data.get('volatility_250d'),
            data.get('annual_return_20d'), data.get('annual_return_60d'),
            data.get('annual_return_120d'), data.get('annual_return_250d'),
            data.get('period_return_20d'), data.get('period_return_60d'),
            data.get('period_return_120d'), data.get('period_return_250d'),
            data.get('ma_20'), data.get('ma_30'), data.get('ma_60'),
            data.get('ma_120'), data.get('ma_250'),
            data.get('win_rate_20d'), data.get('win_rate_60d'),
            data.get('win_rate_120d'), data.get('win_rate_250d'),
            data.get('total_days'), data.get('drift'), data.get('volatility'),
            price_series, market_turnover,
            data.get('data_source', 'tushare_realtime'),
            data.get('generated_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        ))
        conn.commit()
        conn.close()

    # ==================== industry_data ====================

    def load_industry_data(self, stock_code):
        conn = self.get_connection()
        row = conn.execute("SELECT * FROM industry_data WHERE stock_code=%s", (stock_code,)).fetchone()
        conn.close()
        return row if row else None

    def save_industry_data(self, stock_code, data):
        conn = self.get_connection()
        # 同时更新股票的行业分类信息
        self.upsert_stock(stock_code,
                          sw_l1_code=data.get('sw_l1_code'), sw_l1_name=data.get('sw_l1_name'),
                          sw_l2_code=data.get('sw_l2_code'), sw_l2_name=data.get('sw_l2_name'),
                          sw_l3_code=data.get('sw_l3_code'), sw_l3_name=data.get('sw_l3_name'))

        conn.execute("""
            INSERT INTO industry_data (
                stock_code, index_code, industry_name,
                sw_l1_code, sw_l1_name, sw_l2_code, sw_l2_name, sw_l3_code, sw_l3_name,
                analysis_date, current_level,
                volatility_20d, volatility_60d, volatility_120d, volatility_250d,
                annual_return_20d, annual_return_60d, annual_return_120d, annual_return_250d,
                period_return_20d, period_return_60d, period_return_120d, period_return_250d,
                ma_20, ma_60, ma_120, ma_250,
                win_rate_20d, win_rate_60d, win_rate_120d, win_rate_250d,
                total_days, drift, volatility,
                data_source, generated_at
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON DUPLICATE KEY UPDATE
                index_code=VALUES(index_code), industry_name=VALUES(industry_name),
                sw_l1_code=VALUES(sw_l1_code), sw_l1_name=VALUES(sw_l1_name),
                sw_l2_code=VALUES(sw_l2_code), sw_l2_name=VALUES(sw_l2_name),
                sw_l3_code=VALUES(sw_l3_code), sw_l3_name=VALUES(sw_l3_name),
                analysis_date=VALUES(analysis_date), current_level=VALUES(current_level),
                volatility_20d=VALUES(volatility_20d), volatility_60d=VALUES(volatility_60d),
                volatility_120d=VALUES(volatility_120d), volatility_250d=VALUES(volatility_250d),
                annual_return_20d=VALUES(annual_return_20d), annual_return_60d=VALUES(annual_return_60d),
                annual_return_120d=VALUES(annual_return_120d), annual_return_250d=VALUES(annual_return_250d),
                period_return_20d=VALUES(period_return_20d), period_return_60d=VALUES(period_return_60d),
                period_return_120d=VALUES(period_return_120d), period_return_250d=VALUES(period_return_250d),
                ma_20=VALUES(ma_20), ma_60=VALUES(ma_60), ma_120=VALUES(ma_120), ma_250=VALUES(ma_250),
                win_rate_20d=VALUES(win_rate_20d), win_rate_60d=VALUES(win_rate_60d),
                win_rate_120d=VALUES(win_rate_120d), win_rate_250d=VALUES(win_rate_250d),
                total_days=VALUES(total_days), drift=VALUES(drift), volatility=VALUES(volatility),
                data_source=VALUES(data_source), generated_at=VALUES(generated_at)
        """, (
            stock_code, data.get('index_code'), data.get('industry_name'),
            data.get('sw_l1_code'), data.get('sw_l1_name'),
            data.get('sw_l2_code'), data.get('sw_l2_name'),
            data.get('sw_l3_code'), data.get('sw_l3_name'),
            data.get('analysis_date'), data.get('current_level'),
            data.get('volatility_20d'), data.get('volatility_60d'),
            data.get('volatility_120d'), data.get('volatility_250d'),
            data.get('annual_return_20d'), data.get('annual_return_60d'),
            data.get('annual_return_120d'), data.get('annual_return_250d'),
            data.get('period_return_20d'), data.get('period_return_60d'),
            data.get('period_return_120d'), data.get('period_return_250d'),
            data.get('ma_20'), data.get('ma_60'), data.get('ma_120'), data.get('ma_250'),
            data.get('win_rate_20d'), data.get('win_rate_60d'),
            data.get('win_rate_120d'), data.get('win_rate_250d'),
            data.get('total_days'), data.get('drift'), data.get('volatility'),
            data.get('data_source', 'tushare_sw_index'),
            data.get('generated_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        ))
        conn.commit()
        conn.close()

    def save_industry_daily(self, index_code, df, data_source='tushare_sw'):
        """保存行业指数日线数据（含PE/PB）到DB

        Args:
            index_code: 行业指数代码
            df: 日线DataFrame
            data_source: 数据来源 'tushare_sw'(申万) 或 'akshare_ths'(同花顺)
        """
        if df is None or df.empty:
            return
        conn = self.get_connection()
        for _, row in df.iterrows():
            conn.execute("""
                INSERT OR REPLACE INTO industry_daily
                    (index_code, trade_date, open, high, low, close, volume, amount, pct_chg, pe, pb, ps_ttm, data_source)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                index_code,
                str(row.get('trade_date', '')),
                row.get('open'), row.get('high'), row.get('low'), row.get('close'),
                row.get('vol', row.get('volume')),
                row.get('amount'),
                row.get('pct_chg', row.get('pct_change')),
                row.get('pe'), row.get('pb'), row.get('ps_ttm'),
                data_source,
            ))
        conn.commit()
        conn.close()

    def load_industry_daily(self, index_code, start_date=None, end_date=None, data_source=None):
        """从DB加载行业指数日线数据

        Args:
            index_code: 行业指数代码
            start_date: 开始日期
            end_date: 结束日期
            data_source: 数据来源过滤，None=不限（优先tushare_sw）
        """
        conn = self.get_connection()
        if data_source:
            if start_date and end_date:
                rows = conn.execute(
                    "SELECT * FROM industry_daily WHERE index_code=%s AND data_source=%s AND trade_date>=%s AND trade_date<=%s ORDER BY trade_date",
                    (index_code, data_source, start_date, end_date)
                ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM industry_daily WHERE index_code=%s AND data_source=%s ORDER BY trade_date",
                    (index_code, data_source)
                ).fetchall()
        else:
            # 优先取申万数据(tushare_sw)，无数据时再取同花顺(akshare_ths)
            if start_date and end_date:
                rows = conn.execute(
                    "SELECT * FROM industry_daily WHERE index_code=%s AND data_source='tushare_sw' AND trade_date>=%s AND trade_date<=%s ORDER BY trade_date",
                    (index_code, start_date, end_date)
                ).fetchall()
                if not rows:
                    rows = conn.execute(
                        "SELECT * FROM industry_daily WHERE index_code=%s AND data_source='akshare_ths' AND trade_date>=%s AND trade_date<=%s ORDER BY trade_date",
                        (index_code, start_date, end_date)
                    ).fetchall()
            else:
                rows = conn.execute(
                    "SELECT * FROM industry_daily WHERE index_code=%s AND data_source='tushare_sw' ORDER BY trade_date",
                    (index_code,)
                ).fetchall()
                if not rows:
                    rows = conn.execute(
                        "SELECT * FROM industry_daily WHERE index_code=%s AND data_source='akshare_ths' ORDER BY trade_date",
                        (index_code,)
                    ).fetchall()
        conn.close()
        if not rows:
            return None
        import pandas as pd
        return pd.DataFrame([r for r in rows])

    def is_industry_data_stale(self, stock_code, max_days=7):
        """检查行业数据是否过期。"""
        conn = self.get_connection()
        row = conn.execute("SELECT generated_at FROM industry_data WHERE stock_code=%s", (stock_code,)).fetchone()
        conn.close()
        if row is None:
            return True
        try:
            gen_time = datetime.strptime(row['generated_at'][:10], '%Y-%m-%d')
            return (datetime.now() - gen_time).days > max_days
        except (ValueError, TypeError):
            return True

    # ==================== relative_valuation ====================

    def is_cache_valid(self, stock_code):
        """检查相对估值缓存是否有效（cache_date == 今天）。"""
        conn = self.get_connection()
        row = conn.execute(
            "SELECT cache_date FROM relative_valuation WHERE stock_code=%s", (stock_code,)
        ).fetchone()
        conn.close()
        if row is None:
            return False
        today = datetime.now().strftime('%Y%m%d')
        return row['cache_date'] == today

    def load_relative_valuation(self, stock_code):
        """加载相对估值缓存（含 peer_companies）。"""
        conn = self.get_connection()
        row = conn.execute(
            "SELECT * FROM relative_valuation WHERE stock_code=%s", (stock_code,)
        ).fetchone()
        if row is None:
            conn.close()
            return None

        result = row
        result.pop('created_at', None)

        # 加载 peer_companies
        peers = conn.execute(
            "SELECT peer_name as name, peer_code as code, pe, ps, pb, market_cap "
            "FROM peer_companies WHERE stock_code=%s", (stock_code,)
        ).fetchall()
        result['peer_companies'] = [dict(p) for p in peers]
        result['current_metrics'] = {
            'pe': result.pop('current_pe'),
            'pb': result.pop('current_pb'),
            'ps': result.pop('current_ps'),
        }
        conn.close()
        return result

    def save_relative_valuation(self, stock_code, data):
        """保存相对估值缓存。"""
        conn = self.get_connection()
        self.upsert_stock(stock_code)

        metrics = data.get('current_metrics', {})
        conn.execute("""
            INSERT INTO relative_valuation (
                stock_code, cache_date, trade_date,
                current_pe, current_pb, current_ps,
                sw_index_pe, sw_index_pb, sw_index_ps,
                target_index_code, target_industry_l3
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            ON DUPLICATE KEY UPDATE
                cache_date=VALUES(cache_date), trade_date=VALUES(trade_date),
                current_pe=VALUES(current_pe), current_pb=VALUES(current_pb),
                current_ps=VALUES(current_ps),
                sw_index_pe=VALUES(sw_index_pe), sw_index_pb=VALUES(sw_index_pb),
                sw_index_ps=VALUES(sw_index_ps),
                target_index_code=VALUES(target_index_code),
                target_industry_l3=VALUES(target_industry_l3),
                created_at=NOW()
        """, (
            stock_code, data.get('cache_date'), data.get('trade_date'),
            metrics.get('pe'), metrics.get('pb'), metrics.get('ps'),
            data.get('sw_index_pe'), data.get('sw_index_pb'), data.get('sw_index_ps'),
            data.get('target_index_code'), data.get('target_industry_l3'),
        ))

        # 清除旧 peer_companies 并重新写入
        conn.execute("DELETE FROM peer_companies WHERE stock_code=%s", (stock_code,))
        for p in data.get('peer_companies', []):
            conn.execute("""
                INSERT INTO peer_companies (stock_code, peer_name, peer_code, pe, ps, pb, market_cap)
                VALUES (%s,%s,%s,%s,%s,%s,%s)
            """, (stock_code, p.get('name'), p.get('code'),
                  p.get('pe'), p.get('ps'), p.get('pb'), p.get('market_cap')))

        conn.commit()
        conn.close()

    # ==================== issue_date_locked ====================

    def load_issue_date_locked(self, stock_code, issue_date):
        conn = self.get_connection()
        row = conn.execute(
            "SELECT * FROM issue_date_locked WHERE stock_code=%s AND issue_date=%s",
            (stock_code, issue_date),
        ).fetchone()
        conn.close()
        return row if row else None

    def save_issue_date_locked(self, stock_code, issue_date, data):
        conn = self.get_connection()
        self.upsert_stock(stock_code)
        conn.execute("""
            INSERT INTO issue_date_locked (stock_code, issue_date, issue_date_price, ma_20,
                current_price, analysis_date, locked_timestamp)
            VALUES (%s,%s,%s,%s,%s,%s,%s)
            ON DUPLICATE KEY UPDATE
                issue_date_price=VALUES(issue_date_price),
                ma_20=VALUES(ma_20),
                current_price=VALUES(current_price),
                analysis_date=VALUES(analysis_date),
                locked_timestamp=VALUES(locked_timestamp)
        """, (
            stock_code, issue_date, data.get('issue_date_price'),
            data.get('ma_20'), data.get('current_price'),
            data.get('analysis_date'), data.get('locked_timestamp'),
        ))
        conn.commit()
        conn.close()

    def update_issue_date_current_price(self, stock_code, issue_date, price, analysis_date):
        conn = self.get_connection()
        conn.execute("""
            UPDATE issue_date_locked SET current_price=%s, analysis_date=%s
            WHERE stock_code=%s AND issue_date=%s
        """, (price, analysis_date, stock_code, issue_date))
        conn.commit()
        conn.close()

    # ==================== market_indices ====================

    def load_market_indices(self, locked_date=None):
        """加载市场指数数据。返回 dict，key 为指数名称。"""
        conn = self.get_connection()
        rows = conn.execute(
            "SELECT * FROM market_indices WHERE locked_date=%s",
            (locked_date or '',),
        ).fetchall()
        conn.close()

        result = {}
        for row in rows:
            r = row
            name = r.pop('index_name')
            r.pop('index_code', None)
            r.pop('locked_date', None)
            r.pop('created_at', None)
            result[name] = r
        return result if result else None

    def save_market_indices(self, indices_data, locked_date=None):
        """保存市场指数数据。indices_data: dict keyed by index_name."""
        conn = self.get_connection()
        ld = locked_date or ''
        # 先清除同 locked_date 的旧数据
        conn.execute("DELETE FROM market_indices WHERE locked_date=%s", (ld,))

        for name, data in indices_data.items():
            # 用 index_name 作为 index_code（JSON 数据中无 index_code 字段）
            idx_code = data.get('index_code') or name
            conn.execute("""
                INSERT OR REPLACE INTO market_indices (
                    index_code, index_name, locked_date, current_level,
                    volatility_20d, volatility_60d, volatility_120d, volatility_250d,
                    return_20d, return_60d, return_120d, return_250d,
                    period_log_return_20d, period_log_return_60d,
                    period_log_return_120d, period_log_return_250d,
                    ma_20, ma_60, ma_120, ma_250,
                    win_rate_20d, win_rate_60d, win_rate_120d, win_rate_250d,
                    data_date
                ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (
                idx_code, name, ld, data.get('current_level'),
                data.get('volatility_20d'), data.get('volatility_60d'),
                data.get('volatility_120d'), data.get('volatility_250d'),
                data.get('return_20d'), data.get('return_60d'),
                data.get('return_120d'), data.get('return_250d'),
                data.get('period_log_return_20d'), data.get('period_log_return_60d'),
                data.get('period_log_return_120d'), data.get('period_log_return_250d'),
                data.get('ma_20'), data.get('ma_60'), data.get('ma_120'), data.get('ma_250'),
                data.get('win_rate_20d'), data.get('win_rate_60d'),
                data.get('win_rate_120d'), data.get('win_rate_250d'),
                data.get('data_date'),
            ))

        conn.commit()
        conn.close()

    # ==================== screening_results ====================

    def save_screening_result(self, batch_id, stock_code, stock_name, result):
        conn = self.get_connection()
        self.upsert_stock(stock_code, stock_name)
        decision = result.get('decision_conclusion') or {}
        pr = decision.get('premium_range') or {}
        conn.execute("""
            INSERT INTO screening_results (
                batch_id, stock_code, stock_name,
                premium_min, premium_max, valid_thresholds,
                step1_pass, step1_detail,
                step2_pass, step2_detail,
                step3_pass, step3_detail,
                decision, summary, error
            ) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (
            batch_id, stock_code, stock_name,
            pr.get('min'), pr.get('max'), decision.get('valid_thresholds'),
            1 if decision.get('step1', {}).get('pass') else 0,
            decision.get('step1', {}).get('detail', ''),
            1 if decision.get('step2', {}).get('pass') else 0,
            decision.get('step2', {}).get('detail', ''),
            1 if decision.get('step3', {}).get('pass') else 0,
            decision.get('step3', {}).get('detail', ''),
            decision.get('decision', ''),
            decision.get('summary', ''),
            result.get('error', ''),
        ))
        conn.commit()
        conn.close()

    def get_screening_results(self, batch_id):
        conn = self.get_connection()
        rows = conn.execute(
            "SELECT * FROM screening_results WHERE batch_id=%s ORDER BY id",
            (batch_id,),
        ).fetchall()
        conn.close()
        return [r for r in rows]

    def list_batches(self):
        conn = self.get_connection()
        rows = conn.execute("""
            SELECT batch_id, COUNT(*) as cnt,
                   SUM(CASE WHEN decision='建议参与本次定向增发' THEN 1 ELSE 0 END) as pass_cnt,
                   MIN(created_at) as created_at
            FROM screening_results
            GROUP BY batch_id
            ORDER BY created_at DESC
        """).fetchall()
        conn.close()
        return [r for r in rows]

    # ==================== 迁移工具 ====================

    def migrate_from_json(self, data_dir):
        """从现有 JSON 文件导入数据到数据库。"""
        import glob

        # placement_params
        for f in sorted(glob.glob(os.path.join(data_dir, '*_placement_params.json'))):
            basename = os.path.basename(f)
            code = basename.replace('_placement_params.json', '').replace('_', '.', 1)
            with open(f, 'r', encoding='utf-8') as fh:
                params = json.load(fh)
            print(f"  迁移 placement_params: {code}")
            self.upsert_stock(code)
            self.save_placement_params(code, params)

        # market_data
        for f in sorted(glob.glob(os.path.join(data_dir, '*_market_data.json'))):
            basename = os.path.basename(f)
            code = basename.replace('_market_data.json', '').replace('_', '.', 1)
            with open(f, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            print(f"  迁移 market_data: {code}")
            self.save_market_data(code, data)

        # industry_data
        for f in sorted(glob.glob(os.path.join(data_dir, '*_industry_data.json'))):
            if '_backup_' in f:
                continue
            basename = os.path.basename(f)
            code = basename.replace('_industry_data.json', '').replace('_', '.', 1)
            with open(f, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            print(f"  迁移 industry_data: {code}")
            self.save_industry_data(code, data)

        # relative_valuation
        for f in sorted(glob.glob(os.path.join(data_dir, '*_relative_valuation.json'))):
            basename = os.path.basename(f)
            code = basename.replace('_relative_valuation.json', '').replace('_', '.', 1)
            with open(f, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            print(f"  迁移 relative_valuation: {code}")
            self.save_relative_valuation(code, data)

        # issue_date_locked
        for f in sorted(glob.glob(os.path.join(data_dir, '*_issue_date_locked.json'))):
            basename = os.path.basename(f)
            code = basename.replace('_issue_date_locked.json', '').replace('_', '.', 1)
            with open(f, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            issue_date = data.get('issue_date', '')
            if issue_date:
                print(f"  迁移 issue_date_locked: {code} @ {issue_date}")
                self.save_issue_date_locked(code, issue_date, data)

        # market_indices_scenario_data_v2.json
        indices_file = os.path.join(data_dir, 'market_indices_scenario_data_v2.json')
        if os.path.exists(indices_file):
            with open(indices_file, 'r', encoding='utf-8') as fh:
                data = json.load(fh)
            print(f"  迁移 market_indices: {len(data)} 个指数")
            self.save_market_indices(data)

        print("\n迁移完成！")

    # ==================== 东方财富行业板块 ====================

    def save_em_industry_boards(self, boards_df):
        """批量保存/更新东方财富行业板块列表。

        :param boards_df: akshare stock_board_industry_name_em() 返回的 DataFrame
        """
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = self.get_connection()
        try:
            for _, row in boards_df.iterrows():
                conn.execute("""
                    INSERT INTO em_industry_boards
                        (board_code, board_name, total_count, latest_price, change_pct,
                         total_mv, turnover_rate, up_count, down_count,
                         leading_stock, leading_pct, updated_at)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                        board_name=VALUES(board_name),
                        total_count=VALUES(total_count),
                        latest_price=VALUES(latest_price),
                        change_pct=VALUES(change_pct),
                        total_mv=VALUES(total_mv),
                        turnover_rate=VALUES(turnover_rate),
                        up_count=VALUES(up_count),
                        down_count=VALUES(down_count),
                        leading_stock=VALUES(leading_stock),
                        leading_pct=VALUES(leading_pct),
                        updated_at=VALUES(updated_at)
                """, (
                    row.get('板块代码'), row.get('板块名称'), 0,
                    row.get('最新价'), row.get('涨跌幅'),
                    row.get('总市值'), row.get('换手率'),
                    row.get('上涨家数'), row.get('下跌家数'),
                    row.get('领涨股票'), row.get('领涨股票-涨跌幅'),
                    now,
                ))
            conn.commit()
        finally:
            conn.close()

    def save_em_industry_stocks(self, board_code, cons_df):
        """保存某个行业板块的成份股。

        :param board_code: 板块代码 (如 BK1027)
        :param cons_df: akshare stock_board_industry_cons_em() 返回的 DataFrame
        """
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        conn = self.get_connection()
        try:
            # 先删除该板块旧的成份股
            conn.execute(
                "DELETE FROM em_industry_stocks WHERE board_code = %s",
                (board_code,)
            )
            for _, row in cons_df.iterrows():
                conn.execute("""
                    INSERT INTO em_industry_stocks
                        (board_code, stock_code, stock_name,
                         latest_price, change_pct, change_amt,
                         volume, amount, amplitude,
                         high, low, open, prev_close,
                         turnover_rate, pe_dynamic, pb, created_at)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    board_code,
                    row.get('代码'), row.get('名称'),
                    row.get('最新价'), row.get('涨跌幅'), row.get('涨跌额'),
                    row.get('成交量'), row.get('成交额'), row.get('振幅'),
                    row.get('最高'), row.get('最低'), row.get('今开'), row.get('昨收'),
                    row.get('换手率'), row.get('市盈率-动态'), row.get('市净率'),
                    now,
                ))
            # 更新板块的成份股数量
            conn.execute(
                "UPDATE em_industry_boards SET total_count = %s, updated_at = %s WHERE board_code = %s",
                (len(cons_df), now, board_code)
            )
            conn.commit()
        finally:
            conn.close()

    def get_em_industry_boards(self):
        """获取所有东方财富行业板块。"""
        conn = self.get_connection()
        try:
            cur = conn.execute(
                "SELECT board_code, board_name, total_count FROM em_industry_boards ORDER BY board_code"
            )
            return [row for row in cur.fetchall()]
        finally:
            conn.close()

    def get_em_industry_stocks(self, board_code):
        """获取某个行业板块的所有成份股。

        :param board_code: 板块代码 (如 BK1027)
        :return: list of dict
        """
        conn = self.get_connection()
        try:
            cur = conn.execute(
                "SELECT * FROM em_industry_stocks WHERE board_code = %s ORDER BY stock_code",
                (board_code,)
            )
            return [row for row in cur.fetchall()]
        finally:
            conn.close()

    def get_stock_industries(self, stock_code):
        """根据股票代码查询其所属的所有东方财富行业板块。

        :param stock_code: 股票代码 (如 600519)
        :return: list of dict with board_code, board_name
        """
        conn = self.get_connection()
        try:
            cur = conn.execute("""
                SELECT b.board_code, b.board_name
                FROM em_industry_stocks s
                JOIN em_industry_boards b ON s.board_code = b.board_code
                WHERE s.stock_code = %s
                ORDER BY b.board_code
            """, (stock_code,))
            return [row for row in cur.fetchall()]
        finally:
            conn.close()

    def get_em_industry_boards_count(self):
        """获取已入库的行业板块和成份股统计。"""
        conn = self.get_connection()
        try:
            cur = conn.execute("""
                SELECT
                    (SELECT COUNT(*) FROM em_industry_boards) AS board_count,
                    (SELECT COUNT(*) FROM em_industry_stocks) AS stock_count,
                    (SELECT COUNT(DISTINCT stock_code) FROM em_industry_stocks) AS unique_stock_count,
                    (SELECT COUNT(*) FROM em_industry_boards WHERE total_count > 0) AS boards_with_stocks
            """)
            return dict(cur.fetchone())
        finally:
            conn.close()


# ==================== Schema ====================

SCHEMA = """
CREATE TABLE IF NOT EXISTS stocks (
    stock_code   TEXT PRIMARY KEY,
    stock_name   TEXT,
    sw_l1_code   TEXT, sw_l1_name   TEXT,
    sw_l2_code   TEXT, sw_l2_name   TEXT,
    sw_l3_code   TEXT, sw_l3_name   TEXT,
    created_at   TEXT DEFAULT (NOW()),
    updated_at   TEXT DEFAULT (NOW())
);

CREATE TABLE IF NOT EXISTS placement_params (
    stock_code       TEXT PRIMARY KEY REFERENCES stocks(stock_code),
    financing_amount INTEGER DEFAULT 100000000,
    lockup_period    INTEGER DEFAULT 6,
    pricing_method   TEXT DEFAULT 'ma20_discount_90',
    premium_rate     REAL DEFAULT -0.10,
    risk_free_rate   REAL DEFAULT 0.03,
    net_assets       REAL DEFAULT 0,
    total_debt       REAL DEFAULT 0,
    net_income       REAL DEFAULT 0,
    revenue_growth   REAL DEFAULT 0.15,
    operating_margin REAL DEFAULT 0.15,
    beta             REAL DEFAULT 1.0,
    created_at       TEXT DEFAULT (NOW()),
    updated_at       TEXT DEFAULT (NOW())
);

CREATE TABLE IF NOT EXISTS historical_fcf (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    stock_code  TEXT NOT NULL REFERENCES stocks(stock_code),
    year        INTEGER NOT NULL,
    revenue     REAL, operate_profit REAL, net_income REAL,
    nopat       REAL, depreciation  REAL, capex      REAL,
    wc_change   REAL, fcf           REAL,
    UNIQUE(stock_code, year)
);

CREATE TABLE IF NOT EXISTS market_data (
    stock_code       TEXT PRIMARY KEY REFERENCES stocks(stock_code),
    analysis_date    TEXT,
    latest_trading_date TEXT,
    issue_date       TEXT, invitation_date TEXT,
    current_price    REAL, avg_price_all REAL, median_price REAL, price_std REAL,
    volatility_20d   REAL, volatility_60d  REAL, volatility_120d REAL, volatility_250d REAL,
    annual_return_20d REAL, annual_return_60d REAL, annual_return_120d REAL, annual_return_250d REAL,
    period_return_20d REAL, period_return_60d REAL, period_return_120d REAL, period_return_250d REAL,
    ma_20 REAL, ma_30 REAL, ma_60 REAL, ma_120 REAL, ma_250 REAL,
    win_rate_20d REAL, win_rate_60d REAL, win_rate_120d REAL, win_rate_250d REAL,
    total_days INTEGER,
    drift      REAL, volatility REAL,
    price_series     TEXT,
    market_turnover  TEXT,
    data_source      TEXT DEFAULT 'tushare_realtime',
    generated_at     TEXT DEFAULT (NOW())
);

CREATE TABLE IF NOT EXISTS industry_data (
    stock_code    TEXT PRIMARY KEY REFERENCES stocks(stock_code),
    index_code    TEXT,
    industry_name TEXT,
    sw_l1_code TEXT, sw_l1_name TEXT,
    sw_l2_code TEXT, sw_l2_name TEXT,
    sw_l3_code TEXT, sw_l3_name TEXT,
    analysis_date  TEXT, current_level REAL,
    volatility_20d REAL, volatility_60d REAL, volatility_120d REAL, volatility_250d REAL,
    annual_return_20d REAL, annual_return_60d REAL, annual_return_120d REAL, annual_return_250d REAL,
    period_return_20d REAL, period_return_60d REAL, period_return_120d REAL, period_return_250d REAL,
    ma_20 REAL, ma_60 REAL, ma_120 REAL, ma_250 REAL,
    win_rate_20d REAL, win_rate_60d REAL, win_rate_120d REAL, win_rate_250d REAL,
    total_days INTEGER, drift REAL, volatility REAL,
    data_source TEXT DEFAULT 'tushare_sw_index',
    generated_at TEXT DEFAULT (NOW())
);

CREATE TABLE IF NOT EXISTS relative_valuation (
    stock_code        TEXT PRIMARY KEY REFERENCES stocks(stock_code),
    cache_date        TEXT NOT NULL,
    trade_date        TEXT,
    current_pe        REAL, current_pb REAL, current_ps REAL,
    sw_index_pe       REAL, sw_index_pb REAL, sw_index_ps REAL,
    target_index_code TEXT, target_industry_l3 TEXT,
    created_at        TEXT DEFAULT (NOW())
);

CREATE TABLE IF NOT EXISTS peer_companies (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    stock_code  TEXT NOT NULL REFERENCES stocks(stock_code),
    peer_name   TEXT, peer_code TEXT,
    pe REAL, ps REAL, pb REAL, market_cap REAL
);

CREATE TABLE IF NOT EXISTS issue_date_locked (
    stock_code      TEXT NOT NULL REFERENCES stocks(stock_code),
    issue_date      TEXT NOT NULL,
    issue_date_price REAL,
    ma_20           REAL,
    current_price   REAL,
    analysis_date   TEXT,
    locked_timestamp TEXT,
    PRIMARY KEY (stock_code, issue_date)
);

CREATE TABLE IF NOT EXISTS market_indices (
    index_code  TEXT NOT NULL,
    index_name  TEXT NOT NULL,
    locked_date TEXT DEFAULT '',
    current_level REAL,
    volatility_20d REAL, volatility_60d REAL, volatility_120d REAL, volatility_250d REAL,
    return_20d REAL, return_60d REAL, return_120d REAL, return_250d REAL,
    period_log_return_20d REAL, period_log_return_60d REAL,
    period_log_return_120d REAL, period_log_return_250d REAL,
    ma_20 REAL, ma_60 REAL, ma_120 REAL, ma_250 REAL,
    win_rate_20d REAL, win_rate_60d REAL, win_rate_120d REAL, win_rate_250d REAL,
    data_date   TEXT,
    created_at  TEXT DEFAULT (NOW()),
    PRIMARY KEY (index_code, locked_date)
);

CREATE TABLE IF NOT EXISTS screening_results (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id        TEXT NOT NULL,
    stock_code      TEXT REFERENCES stocks(stock_code),
    stock_name      TEXT,
    premium_min     REAL, premium_max   REAL,
    valid_thresholds INTEGER,
    step1_pass      INTEGER, step1_detail TEXT,
    step2_pass      INTEGER, step2_detail TEXT,
    step3_pass      INTEGER, step3_detail TEXT,
    decision        TEXT,
    summary         TEXT,
    error           TEXT,
    created_at      TEXT DEFAULT (NOW())
);
CREATE INDEX IF NOT EXISTS idx_screening_batch ON screening_results(batch_id);

CREATE TABLE IF NOT EXISTS industry_daily (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    index_code      TEXT NOT NULL,
    trade_date      TEXT NOT NULL,
    open            REAL,
    high            REAL,
    low             REAL,
    close           REAL,
    volume          REAL,
    amount          REAL,
    pct_chg         REAL,
    pe              REAL,
    pb              REAL,
    ps_ttm          REAL,
    data_source     TEXT DEFAULT 'tushare_sw',
    UNIQUE(index_code, trade_date, data_source)
);
CREATE INDEX IF NOT EXISTS idx_industry_daily_code ON industry_daily(index_code);
CREATE INDEX IF NOT EXISTS idx_industry_daily_date ON industry_daily(index_code, trade_date);

-- 东方财富行业板块列表
CREATE TABLE IF NOT EXISTS em_industry_boards (
    board_code     TEXT PRIMARY KEY,  -- 板块代码 (BK1027)
    board_name     TEXT NOT NULL,     -- 板块名称 (小金属)
    total_count    INTEGER DEFAULT 0,-- 成份股数量
    latest_price   REAL,             -- 最新价
    change_pct     REAL,             -- 涨跌幅
    total_mv       REAL,             -- 总市值
    turnover_rate  REAL,             -- 换手率
    up_count       INTEGER,          -- 上涨家数
    down_count     INTEGER,          -- 下跌家数
    leading_stock  TEXT,             -- 领涨股票
    leading_pct    REAL,             -- 领涨股票涨跌幅
    created_at     TEXT DEFAULT (NOW()),
    updated_at     TEXT DEFAULT (NOW())
);

-- 东方财富行业板块成份股
CREATE TABLE IF NOT EXISTS em_industry_stocks (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    board_code     TEXT NOT NULL,     -- 所属板块代码
    stock_code     TEXT NOT NULL,     -- 股票代码
    stock_name     TEXT NOT NULL,     -- 股票名称
    latest_price   REAL,             -- 最新价
    change_pct     REAL,             -- 涨跌幅
    change_amt     REAL,             -- 涨跌额
    volume         REAL,             -- 成交量
    amount         REAL,             -- 成交额
    amplitude      REAL,             -- 振幅
    high           REAL,             -- 最高
    low            REAL,             -- 最低
    open           REAL,             -- 今开
    prev_close     REAL,             -- 昨收
    turnover_rate  REAL,             -- 换手率
    pe_dynamic     REAL,             -- 市盈率-动态
    pb             REAL,             -- 市净率
    created_at     TEXT DEFAULT (NOW()),
    UNIQUE(board_code, stock_code)
);
CREATE INDEX IF NOT EXISTS idx_em_industry_stocks_board ON em_industry_stocks(board_code);
CREATE INDEX IF NOT EXISTS idx_em_industry_stocks_code ON em_industry_stocks(stock_code);
"""
