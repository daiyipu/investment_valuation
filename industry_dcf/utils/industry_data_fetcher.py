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

        # Fetch data for each company
        companies = {}
        start_date = (datetime.now() - timedelta(days=years * 365)).strftime('%Y%m%d')

        for i, ts_code in enumerate(member_codes):
            if (i + 1) % 10 == 0:
                print(f"    进度: {i+1}/{len(member_codes)}")
            data = self._fetch_company_data(ts_code, start_date)
            if data is not None:
                companies[ts_code] = data

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
    ) -> Optional[Dict[str, pd.DataFrame]]:
        """Fetch financial data for a single company as DataFrames.

        Returns {'cashflow': DataFrame, 'income': DataFrame, 'balance': DataFrame}
        """
        start_date = (datetime.now() - timedelta(days=years * 365)).strftime('%Y%m%d')
        data = self._fetch_company_data(ts_code, start_date)
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
        """
        if not trade_date:
            trade_date = datetime.now().strftime('%Y%m%d')

        member_codes = get_l3_members(l3_code, self.pro)
        if not member_codes:
            return pd.DataFrame()

        results = []
        for ts_code in member_codes:
            db = self.rate_limiter.call(
                self.pro.daily_basic,
                ts_code=ts_code, trade_date=trade_date,
                fields='ts_code,trade_date,pe,pb,total_mv,total_share',
            )
            if db is not None and not db.empty:
                results.append(db)

        # Try previous dates if no data found
        if not results:
            for days_back in range(1, 10):
                td = (datetime.now() - timedelta(days=days_back)).strftime('%Y%m%d')
                for ts_code in member_codes:
                    db = self.rate_limiter.call(
                        self.pro.daily_basic,
                        ts_code=ts_code, trade_date=td,
                        fields='ts_code,trade_date,pe,pb,total_mv,total_share',
                    )
                    if db is not None and not db.empty:
                        results.append(db)
                if results:
                    break

        if not results:
            return pd.DataFrame()

        valid = [df.dropna(axis=1, how='all') for df in results if not df.empty]
        if not valid:
            return pd.DataFrame()
        return pd.concat(valid, ignore_index=True)

    def _fetch_company_data(self, ts_code: str, start_date: str) -> Optional[dict]:
        """Fetch cashflow, income, balance sheet for one company."""
        try:
            cashflow = self.rate_limiter.call(
                self.pro.cashflow,
                ts_code=ts_code, start_date=start_date,
            )
            income = self.rate_limiter.call(
                self.pro.income,
                ts_code=ts_code, start_date=start_date,
            )
            balance = self.rate_limiter.call(
                self.pro.balancesheet,
                ts_code=ts_code, start_date=start_date,
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
        """Save industry data to JSON cache file."""
        filepath = os.path.join(self.cache_dir, f'{l3_code}_financials.json')
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return filepath

    def _load_cache(self, l3_code: str, max_age_hours: int = 24) -> Optional[dict]:
        """Load cached data if fresh enough. Returns None if missing/stale."""
        filepath = os.path.join(self.cache_dir, f'{l3_code}_financials.json')
        if not os.path.exists(filepath):
            return None

        try:
            # Check file age
            mtime = os.path.getmtime(filepath)
            age_hours = (time.time() - mtime) / 3600
            if age_hours > max_age_hours:
                return None

            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            print(f"  缓存文件损坏: {e}")
            return None
