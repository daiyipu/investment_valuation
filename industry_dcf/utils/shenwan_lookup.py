#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shenwan (SW) industry classification lookup utilities.

Maps stock codes to their Shenwan L3 industry and lists industry members.
Uses shenwan_industries.json + Tushare index_member API.
"""

import json
import os
import re
from typing import Dict, List, Optional

_L3_PATTERN = re.compile(r'^85\d{4}\.SI$')
_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
_DEFAULT_JSON = os.path.join(_PROJECT_ROOT, 'shenwan_industries.json')


def load_shenwan_tree(json_path: str = None) -> list:
    """Load the shenwan_industries.json classification tree."""
    path = json_path or _DEFAULT_JSON
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def get_all_l3_codes(json_path: str = None) -> List[Dict]:
    """Flatten the Shenwan tree to return all L3 entries.

    Returns list of {
        'l1_code', 'l1_name', 'l2_code', 'l2_name', 'l3_code', 'l3_name'
    }
    """
    tree = load_shenwan_tree(json_path)
    result = []
    for l1 in tree:
        for l2 in l1.get('children', []):
            for l3 in l2.get('children', []):
                result.append({
                    'l1_code': l1['code'],
                    'l1_name': l1['name'],
                    'l2_code': l2['code'],
                    'l2_name': l2['name'],
                    'l3_code': l3['code'],
                    'l3_name': l3['name'],
                })
    return result


def find_l3_by_code(ts_code: str, pro_api, json_path: str = None) -> Optional[Dict]:
    """Find a company's Shenwan L3 industry.

    Strategy:
    1. Call pro.index_member(ts_code=ts_code) to get all index memberships.
    2. Filter for active L3 codes (850xxx.SI pattern).
    3. Cross-reference with shenwan_industries.json for L1/L2 names.

    Returns dict with l1_code, l1_name, l2_code, l2_name, l3_code, l3_name,
    or None if not found.
    """
    # Get all L3 codes for name lookup
    all_l3 = get_all_l3_codes(json_path)
    l3_map = {e['l3_code']: e for e in all_l3}

    try:
        members = pro_api.index_member(ts_code=ts_code)
        if members is None or members.empty:
            return _fallback_by_name(ts_code, pro_api, all_l3)

        # Filter for active L3 memberships
        active = members[members['out_date'].isna()] if 'out_date' in members.columns else members
        l3_codes = [c for c in active['index_code'].tolist() if _L3_PATTERN.match(c)]

        if l3_codes:
            l3_code = l3_codes[0]
            if l3_code in l3_map:
                return l3_map[l3_code]
            # L3 code exists but not in our JSON (rare) — return code only
            return {'l3_code': l3_code, 'l3_name': l3_code, 'l2_code': '', 'l2_name': '',
                    'l1_code': '', 'l1_name': ''}

    except Exception as e:
        print(f"  查询行业分类失败: {e}")

    return _fallback_by_name(ts_code, pro_api, all_l3)


def _fallback_by_name(ts_code: str, pro_api, all_l3: list) -> Optional[Dict]:
    """Fallback: match via stock_basic industry name against L3 names."""
    try:
        basic = pro_api.stock_basic(ts_code=ts_code, fields='ts_code,industry')
        if basic is None or basic.empty:
            return None
        industry_name = basic.iloc[0].get('industry', '')
        if not industry_name:
            return None
        # Match against L3 names
        for entry in all_l3:
            if industry_name in entry['l3_name'] or entry['l3_name'] in industry_name:
                return entry
        # Match against L2 names as further fallback
        l2_names = {e['l2_name'] for e in all_l3}
        for e in all_l3:
            if industry_name in e['l2_name']:
                return e
    except Exception:
        pass
    return None


def get_l3_members(l3_code: str, pro_api) -> List[str]:
    """Get all current member stock codes for a Shenwan L3 industry.

    Returns list of ts_code strings.
    """
    try:
        members = pro_api.index_member(index_code=l3_code)
        if members is None or members.empty:
            return []
        if 'out_date' in members.columns:
            members = members[members['out_date'].isna()]
        return members['con_code'].tolist()
    except Exception as e:
        print(f"  获取行业成员失败({l3_code}): {e}")
        return []
