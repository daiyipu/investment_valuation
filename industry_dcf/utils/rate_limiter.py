#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rate limiter for batch Tushare API calls.

Prevents hitting Tushare's per-minute rate limits (200/min for paid users)
by enforcing minimum intervals between calls and providing retry logic.
"""

import time
from typing import Any, Callable, List, Optional


class RateLimiter:
    """Token-bucket style rate limiter for Tushare API calls."""

    def __init__(self, calls_per_minute: int = 180, min_interval: float = 0.35):
        self.calls_per_minute = calls_per_minute
        self.min_interval = min_interval
        self._timestamps: List[float] = []
        self._max_retries = 3

    def _wait_if_needed(self):
        """Sleep if we've exceeded the rate limit."""
        now = time.time()
        # Clean timestamps older than 60s
        self._timestamps = [t for t in self._timestamps if now - t < 60]

        if len(self._timestamps) >= self.calls_per_minute:
            # Wait until the oldest call is 60s ago
            wait = 60 - (now - self._timestamps[0]) + 0.1
            if wait > 0:
                time.sleep(wait)

        # Enforce minimum interval
        if self._timestamps:
            elapsed = now - self._timestamps[-1]
            if elapsed < self.min_interval:
                time.sleep(self.min_interval - elapsed)

    def call(self, func: Callable, *args, **kwargs) -> Any:
        """Execute an API function with rate limiting and retry.

        Returns the result or None on failure.
        """
        for attempt in range(self._max_retries):
            try:
                self._wait_if_needed()
                result = func(*args, **kwargs)
                self._timestamps.append(time.time())
                return result
            except Exception as e:
                if attempt < self._max_retries - 1:
                    wait = 2 ** attempt
                    print(f"    请求失败(第{attempt+1}次重试): {e}, 等待{wait}s")
                    time.sleep(wait)
                else:
                    print(f"    请求失败(已耗尽重试次数): {e}")
                    return None
        return None

    def batch_call(
        self,
        func: Callable,
        param_list: List[dict],
        desc: str = "",
    ) -> List[Any]:
        """Execute the same API function across many parameter sets.

        Args:
            func: API function to call.
            param_list: List of kwargs dicts, one per call.
            desc: Description for progress display.

        Returns:
            List of results (None entries for failures).
        """
        total = len(param_list)
        results = []
        for i, params in enumerate(param_list):
            if desc and (i + 1) % 10 == 0:
                print(f"  {desc} {i+1}/{total}...")
            result = self.call(func, **params)
            results.append(result)
        if desc:
            print(f"  {desc} 完成: {total}条")
        return results
