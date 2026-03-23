# -*- coding: utf-8 -*-
"""
定增风险分析报告生成器（模块化版本）

功能：按章节拆分的报告生成器，降低单个文件复杂度
"""

__version__ = '2.0.0'
__author__ = 'Investment Valuation Team'

from .main import generate_report

__all__ = ['generate_report']
