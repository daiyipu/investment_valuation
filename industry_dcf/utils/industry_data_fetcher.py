#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Industry data fetcher with caching.

Fetches financial statements (cashflow, income, balance sheet) for all
companies in a Shenwan L3 industry, caches locally as JSON files.
"""

import json
import os
import time
from datetime import datetime, timedelta
from typing import Dict, Optional

import pandas as pd

from .rate_limiter import RateLimiter
from .shenwan_lookup import get_l3_members

_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
_DEFAULT_CACHE_DIR = os.path.join(os.path.dirname(__file__), '..', 'data')


class IndustryDataFetcher:
    """Fetch and cache financial data for all companies in a Shenwan L3 industry."""

    def __init__(self, pro_api, cache_dir: str = None, rate_limiter: RateLimiter = None):
        self.pro = pro_api
        self.cache_dir = cache_dir or _DEFAULT_CACHE_DIR
        self.rate_limiter = rate_limiter or RateLimiter()
        os.makedirs(self.cache_dir, exist_ok=True)

    def get_industry_financials(
        self,
        l3_code: str,
        years: int = 10,
        force_refresh: bool = False,
        issue_date: str = None,
    ) -> dict:
        """Get full financial data for an L3 industry.

        Returns:
        {
            'l3_code': str,
            'l3_name': str,
            'fetch_date': str,
            'company_count': int,
            'companies': {
                'ts_code': {
                    'cashflow': list-of-dicts,
                    'income': list-of-dicts,
                    'balance': list-of-dicts,
                },
                ...
            }
        }
        """
        # Try cache first
        if not force_refresh:
            cached = self._load_cache(l3_code, max_age_hours=24)
            if cached is not None:
                print(f"  使用缓存数据: {l3_code} ({cached['company_count']}家公司)")
                return cached

        # Fetch member list
        member_codes = get_l3_members(l3_code, self.pro)
        if not member_codes:
            print(f"  行业 {l3_code} 无成员公司")
            return {'l3_code': l3_code, 'company_count': 0, 'companies': {}}

        print(f"  获取行业 {l3_code} 数据: {len(member_codes)}家公司")

        # Fetch data for each company（并行，避免21家×3报表串行）
        # 回溯分析时以报价日为end_date，避免取报价日后发布的财报
        companies = {}
        ref_dt = datetime.strptime(issue_date, '%Y%m%d') if issue_date else datetime.now()
        start_date = (ref_dt - timedelta(days=years * 365)).strftime('%Y%m%d')
        end_date = ref_dt.strftime('%Y%m%d') if issue_date else None
        from concurrent.futures import ThreadPoolExecutor

        def _fetch_one(code):
            return code, self._fetch_company_data(code, start_date, end_date=end_date)

        with ThreadPoolExecutor(max_workers=8) as executor:
            for code, data in executor.map(_fetch_one, member_codes):
                if data is not None:
                    companies[code] = data
        print(f"    并行获取完成: {len(companies)}/{len(member_codes)} 家")

        result = {
            'l3_code': l3_code,
            'fetch_date': datetime.now().strftime('%Y%m%d'),
            'company_count': len(companies),
            'companies': companies,
        }

        # Save cache
        self._save_cache(l3_code, result)

        return result

    def get_company_financials(
        self,
        ts_code: str,
        years: int = 10,
        issue_date: str = None,
    ) -> Optional[Dict[str, pd.DataFrame]]:
        """Fetch financial data for a single company as DataFrames.

        Returns {'cashflow': DataFrame, 'income': DataFrame, 'balance': DataFrame}
        """
        ref_dt = datetime.strptime(issue_date, '%Y%m%d') if issue_date else datetime.now()
        start_date = (ref_dt - timedelta(days=years * 365)).strftime('%Y%m%d')
        end_date = issue_date  # 限制财报不超过报价日
        data = self._fetch_company_data(ts_code, start_date, end_date=end_date)
        if data is None:
            return None

        return {
            'cashflow': pd.DataFrame(data.get('cashflow') or []),
            'income': pd.DataFrame(data.get('income') or []),
            'balance': pd.DataFrame(data.get('balance') or []),
        }

    def get_industry_daily_basics(
        self,
        l3_code: str,
        trade_date: str = None,
    ) -> pd.DataFrame:
        """Get daily_basic (PE, PB, market cap) for all industry members.

        Used for industry median PE calculation.
        优化：用 daily_basic(trade_date=X) 一次取全市场再筛选，替代逐家循环。
        """
        import pandas as pd

        if not trade_date:
            trade_date = datetime.now().strftime('%Y%m%d')

        member_codes = get_l3_members(l3_code, self.pro)
        if not member_codes:
            return pd.DataFrame()

        member_set = set(member_codes)

        # 一次取全市场 daily_basic，本地筛选行业成员（1次调用替代N次）
        # 回退序列基于传入的trade_date（报价日），而非now()
        base_dt = datetime.strptime(trade_date, '%Y%m%d')
        for td in [trade_date] + [(base_dt - timedelta(days=d)).strftime('%Y%m%d') for d in range(1, 10)]:
            try:
                df_all = self.rate_limiter.call(
                    self.pro.daily_basic,
                    trade_date=td,
                    fields='ts_code,trade_date,pe,pb,total_mv,total_share',
                )
                if df_all is not None and not df_all.empty:
                    df_filtered = df_all[df_all['ts_code'].isin(member_set)].copy()
                    if not df_filtered.empty:
                        return df_filtered.dropna(axis=1, how='all')
            except Exception:
                continue

        # 降级：逐家查询（全市场查询失败时）
        results = []
        for ts_code in member_codes:
            db = self.rate_limiter.call(
                self.pro.daily_basic,
                ts_code=ts_code, trade_date=trade_date,
                fields='ts_code,trade_date,pe,pb,total_mv,total_share',
            )
            if db is not None and not db.empty:
                results.append(db)

        if not results:
            return pd.DataFrame()
        valid = [df.dropna(axis=1, how='all') for df in results if not df.empty]
        return pd.concat(valid, ignore_index=True) if valid else pd.DataFrame()

    def _fetch_company_data(self, ts_code: str, start_date: str, end_date: str = None) -> Optional[dict]:
        """Fetch cashflow, income, balance sheet for one company.

        end_date: 限制财报公告日不超过该日(回溯分析时=报价日)，None则不限制。
        """
        try:
            kw_end = {'end_date': end_date} if end_date else {}
            cashflow = self.rate_limiter.call(
                self.pro.cashflow,
                ts_code=ts_code, start_date=start_date, **kw_end,
            )
            income = self.rate_limiter.call(
                self.pro.income,
                ts_code=ts_code, start_date=start_date, **kw_end,
            )
            balance = self.rate_limiter.call(
                self.pro.balancesheet,
                ts_code=ts_code, start_date=start_date, **kw_end,
            )

            # At least one statement must exist
            has_data = (
                (cashflow is not None and not cashflow.empty)
                or (income is not None and not income.empty)
                or (balance is not None and not balance.empty)
            )
            if not has_data:
                return None

            return {
                'cashflow': cashflow.to_dict('records') if cashflow is not None and not cashflow.empty else [],
                'income': income.to_dict('records') if income is not None and not income.empty else [],
                'balance': balance.to_dict('records') if balance is not None and not balance.empty else [],
            }
        except Exception as e:
            print(f"    获取 {ts_code} 数据失败: {e}")
            return None

    def _save_cache(self, l3_code: str, data: dict) -> str:
        """Save industry data to DB (valuation.db)。"""
        try:
            import sqlite3
            db_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
                os.path.dirname(os.path.abspath(__file__))))),
                'price_maintenance_risk_analysis', 'data', 'valuation.db')
            conn = sqlite3.connect(db_path, timeout=30)
            conn.execute("PRAGMA busy_timeout=30000")
            # 确保表存在
            conn.execute("""
                CREATE TABLE IF NOT EXISTS industry_financials (
                    l3_code TEXT PRIMARY KEY,
                    l3_name TEXT,
                    fetch_date TEXT,
                    company_count INTEGER,
                    data_json TEXT,
                    created_at TEXT
                )
            """)
            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            conn.execute("""
                INSERT OR REPLACE INTO industry_financials
                    (l3_code, l3_name, fetch_date, company_count, data_json, created_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                l3_code,
                data.get('l3_name', ''),
                data.get('fetch_date', ''),
                data.get('company_count', 0),
                json.dumps(data, ensure_ascii=False),
                now
            ))
            conn.commit()
            conn.close()
        except Exception as e:
            print(f"  ⚠️ DB缓存保存失败({l3_code}): {e}，降级到JSON")
            # 降级：写JSON
            filepath = os.path.join(self.cache_dir, f'{l3_code}_financials.json')
            os.makedirs(self.cache_dir, exist_ok=True)
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        return l3_code

    def _load_cache(self, l3_code: str, max_age_hours: int = 24) -> Optional[dict]:
        """Load cached data from DB (valuation.db)。"""
        # 优先从DB读取
        try:
            import sqlite3
            db_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
                os.path.dirname(os.path.abspath(__file__))))),
                'price_maintenance_risk_analysis', 'data', 'valuation.db')
            conn = sqlite3.connect(db_path, timeout=30)
            conn.execute("PRAGMA busy_timeout=30000")
            conn.row_factory = sqlite3.Row
            row = conn.execute(
                "SELECT * FROM industry_financials WHERE l3_code=?", (l3_code,)
            ).fetchone()
            conn.close()

            if row:
                # 检查过期
                created = row['created_at']
                created_dt = datetime.strptime(created[:19], '%Y-%m-%d %H:%M:%S')
                age_hours = (datetime.now() - created_dt).total_seconds() / 3600
                if age_hours > max_age_hours:
                    return None
                return json.loads(row['data_json'])
        except Exception:
            pass

        # 降级：读JSON（兼容旧缓存）
        filepath = os.path.join(self.cache_dir, f'{l3_code}_financials.json')
        if not os.path.exists(filepath):
            return None
        try:
            mtime = os.path.getmtime(filepath)
            age_hours = (time.time() - mtime) / 3600
            if age_hours > max_age_hours:
                return None
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return None
