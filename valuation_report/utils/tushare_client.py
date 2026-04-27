#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tushare API客户端封装
"""

import tushare as ts
from typing import Optional
import time


class TushareClient:
    """Tushare API客户端"""

    def __init__(self, token: str):
        """初始化Tushare客户端

        Args:
            token: Tushare API Token
        """
        self.token = token
        self.pro = None
        self._connect()

    def _connect(self) -> None:
        """建立Tushare连接"""
        try:
            ts.set_token(self.token)
            self.pro = ts.pro_api()
            print("Tushare API连接成功")
        except Exception as e:
            print(f"Tushare API连接失败: {e}")
            raise

    def get_pro(self):
        """获取pro API对象"""
        return self.pro

    def get_stock_basic(self, ts_code: str = None, list_status: str = 'L') -> Optional[any]:
        """获取股票基本信息

        Args:
            ts_code: 股票代码
            list_status: 上市状态 L-上市 D-退市 P-暂停

        Returns:
            股票基本信息DataFrame
        """
        try:
            df = self.pro.stock_basic(
                ts_code=ts_code,
                list_status=list_status,
                fields='ts_code,name,industry,market,list_date,exchange'
            )
            return df
        except Exception as e:
            print(f"获取股票基本信息失败: {e}")
            return None

    def get_daily(self, ts_code: str, start_date: str = None, end_date: str = None) -> Optional[any]:
        """获取日线行情数据

        Args:
            ts_code: 股票代码
            start_date: 开始日期
            end_date: 结束日期

        Returns:
            日线行情数据DataFrame
        """
        try:
            df = self.pro.daily(ts_code=ts_code, start_date=start_date, end_date=end_date)
            return df
        except Exception as e:
            print(f"获取日线行情数据失败: {e}")
            return None

    def get_fina_indicator(self, ts_code: str, start_date: str = None, end_date: str = None) -> Optional[any]:
        """获取财务指标数据

        Args:
            ts_code: 股票代码
            start_date: 开始日期
            end_date: 结束日期

        Returns:
            财务指标数据DataFrame
        """
        try:
            df = self.pro.fina_indicator(ts_code=ts_code, start_date=start_date, end_date=end_date)
            return df
        except Exception as e:
            print(f"获取财务指标数据失败: {e}")
            return None

    def get_income(self, ts_code: str, start_date: str = None, end_date: str = None) -> Optional[any]:
        """获取利润表数据

        Args:
            ts_code: 股票代码
            start_date: 开始日期
            end_date: 结束日期

        Returns:
            利润表数据DataFrame
        """
        try:
            df = self.pro.income(ts_code=ts_code, start_date=start_date, end_date=end_date)
            return df
        except Exception as e:
            print(f"获取利润表数据失败: {e}")
            return None

    def get_balance_sheet(self, ts_code: str, start_date: str = None, end_date: str = None) -> Optional[any]:
        """获取资产负债表数据

        Args:
            ts_code: 股票代码
            start_date: 开始日期
            end_date: 结束日期

        Returns:
            资产负债表数据DataFrame
        """
        try:
            df = self.pro.balancesheet(ts_code=ts_code, start_date=start_date, end_date=end_date)
            return df
        except Exception as e:
            print(f"获取资产负债表数据失败: {e}")
            return None

    def get_cashflow(self, ts_code: str, start_date: str = None, end_date: str = None) -> Optional[any]:
        """获取现金流量表数据

        Args:
            ts_code: 股票代码
            start_date: 开始日期
            end_date: 结束日期

        Returns:
            现金流量表数据DataFrame
        """
        try:
            df = self.pro.cashflow(ts_code=ts_code, start_date=start_date, end_date=end_date)
            return df
        except Exception as e:
            print(f"获取现金流量表数据失败: {e}")
            return None

    def get_daily_basic(self, ts_code: str, start_date: str = None, end_date: str = None) -> Optional[any]:
        """获取每日指标数据

        Args:
            ts_code: 股票代码
            start_date: 开始日期
            end_date: 结束日期

        Returns:
            每日指标数据DataFrame
        """
        try:
            df = self.pro.daily_basic(ts_code=ts_code, start_date=start_date, end_date=end_date)
            return df
        except Exception as e:
            print(f"获取每日指标数据失败: {e}")
            return None

    def get_fina_mainbz(self, ts_code: str, type: str = 'P', start_date: str = None, end_date: str = None) -> Optional[any]:
        """获取主营业务构成

        Args:
            ts_code: 股票代码
            type: 类型 P-按产品 D-按地区
            start_date: 开始日期
            end_date: 结束日期

        Returns:
            主营业务构成数据DataFrame
        """
        try:
            df = self.pro.fina_mainbz(ts_code=ts_code, type=type, start_date=start_date, end_date=end_date)
            return df
        except Exception as e:
            print(f"获取主营业务构成数据失败: {e}")
            return None

    def retry_request(self, func, *args, max_retries: int = 3, delay: float = 1.0, **kwargs):
        """带重试的请求

        Args:
            func: 请求函数
            *args: 函数参数
            max_retries: 最大重试次数
            delay: 重试延迟(秒)
            **kwargs: 函数关键字参数

        Returns:
            请求结果
        """
        for i in range(max_retries):
            try:
                result = func(*args, **kwargs)
                return result
            except Exception as e:
                if i < max_retries - 1:
                    print(f"请求失败，第{i+1}次重试: {e}")
                    time.sleep(delay)
                else:
                    print(f"请求失败，已达到最大重试次数: {e}")
                    return None
