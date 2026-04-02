#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
WACC计算模块
使用CAPM模型计算加权平均资本成本

方案C（混合实现）：
- Beta自动计算（个股和行业）
- 市值/债务从API获取
- 无风险利率、债务成本使用默认值
"""

import numpy as np
import pandas as pd
from typing import Dict, Optional, Tuple
from datetime import datetime, timedelta


class WACCCalculator:
    """WACC计算器"""

    # 默认参数
    DEFAULT_PARAMS = {
        'risk_free_rate': 0.0185,    # Rf: 无风险利率 1.85%（10年期国债收益率，2025年末）
        'market_return': 0.08,       # Rm: 市场收益率 8%（沪深300预期收益率）
        'market_premium': 0.0615,    # Rm - Rf: 市场风险溢价 6.15% (= 8% - 1.85%)
        'debt_premium': 0.50,        # 债务成本上浮比例 50%（相对于无风险利率）
        'tax_rate': 0.25,            # 企业所得税率 25%
        'beta_window': 500,          # Beta计算窗口（交易日，约2年）
    }

    def __init__(self, pro_api=None):
        """
        初始化WACC计算器

        参数:
            pro_api: tushare pro_api对象（如果为None，则不使用API）
        """
        self.pro = pro_api
        self.cached_beta = {}  # 缓存已计算的beta

    def calculate_beta(
        self,
        stock_returns: pd.Series,
        market_returns: pd.Series
    ) -> float:
        """
        计算Beta系数

        参数:
            stock_returns: 个股收益率序列（百分比形式，如2.5表示2.5%）
            market_returns: 市场收益率序列（百分比形式）

        返回:
            beta系数
        """
        # 转换为小数形式
        stock_ret_decimal = stock_returns / 100.0
        market_ret_decimal = market_returns / 100.0

        # 对齐数据长度
        min_len = min(len(stock_ret_decimal), len(market_ret_decimal))
        stock_ret_aligned = stock_ret_decimal.iloc[-min_len:]
        market_ret_aligned = market_ret_decimal.iloc[-min_len:]

        if min_len < 20:
            print(f"⚠️ 数据不足（{min_len}期），无法计算可靠Beta")
            return 1.0

        # 计算协方差和方差
        covariance = np.cov(stock_ret_aligned, market_ret_aligned)[0][1]
        variance = np.var(market_ret_aligned)

        if variance == 0:
            print(f"⚠️ 市场方差为0，无法计算Beta")
            return 1.0

        beta = covariance / variance
        return float(beta)

    def calculate_beta_from_api(
        self,
        stock_code: str,
        window: int = 120,
        market_index: str = '000300.SH'
    ) -> Dict:
        """
        从API获取数据并计算Beta

        参数:
            stock_code: 股票代码（如'300735.SZ'）
            window: 计算窗口（交易日）
            market_index: 市场指数代码（默认沪深300）

        返回:
            {
                'beta': float,
                'stock_code': str,
                'window': int,
                'data_points': int,
                'calculation_date': str
            }
        """
        if self.pro is None:
            return {'beta': 1.0, 'error': 'API未初始化'}

        try:
            # 计算日期范围
            end_date = datetime.now().strftime('%Y%m%d')
            start_date = (datetime.now() - timedelta(days=window*2)).strftime('%Y%m%d')

            # 获取个股数据
            df_stock = self.pro.daily(
                ts_code=stock_code,
                start_date=start_date,
                end_date=end_date,
                fields='trade_date,pct_chg'
            )

            if df_stock.empty:
                return {'beta': 1.0, 'error': '未获取到个股数据'}

            # 获取市场指数数据
            df_market = self.pro.index_daily(
                ts_code=market_index,
                start_date=start_date,
                end_date=end_date,
                fields='trade_date,pct_chg'
            )

            if df_market.empty:
                return {'beta': 1.0, 'error': '未获取到市场数据'}

            # 按日期排序
            df_stock = df_stock.sort_values('trade_date').reset_index(drop=True)
            df_market = df_market.sort_values('trade_date').reset_index(drop=True)

            # 取最近window期数据
            stock_returns = df_stock.iloc[-window:]['pct_chg']
            market_returns = df_market.iloc[-window:]['pct_chg']

            # 计算Beta
            beta = self.calculate_beta(stock_returns, market_returns)

            return {
                'beta': round(beta, 3),
                'stock_code': stock_code,
                'market_index': market_index,
                'window': window,
                'data_points': len(stock_returns),
                'calculation_date': datetime.now().strftime('%Y-%m-%d')
            }

        except Exception as e:
            print(f"⚠️ 计算Beta失败: {e}")
            return {'beta': 1.0, 'error': str(e)}

    def calculate_industry_beta(
        self,
        industry_code: str = None,
        peer_companies: pd.DataFrame = None,
        window: int = 500,
        max_stocks: int = 10
    ) -> Dict:
        """
        计算行业Beta（同行公司Beta的中位数）

        参数:
            industry_code: 申万行业代码（可选，如果提供peer_companies则不需要）
            peer_companies: 同行公司数据（从第二章获取，优先使用）
            window: 计算窗口（默认500天，与个股Beta保持一致）
            max_stocks: 最多计算多少只同行公司

        返回:
            {
                'industry_beta': float,
                'stock_count': int,
                'sample_stocks': list
            }
        """
        if self.pro is None:
            return {'industry_beta': 1.0, 'error': 'API未初始化'}

        stock_codes = []

        # 优先使用第二章的同行公司数据
        if peer_companies is not None and not peer_companies.empty:
            print(f"📊 使用第二章同行公司数据计算行业Beta，同行数量：{len(peer_companies)}")
            # 从同行公司数据中提取股票代码（支持ts_code和code两种列名）
            if 'ts_code' in peer_companies.columns:
                stock_codes = peer_companies['ts_code'].unique().tolist()[:max_stocks]
                print(f"   提取到 {len(stock_codes)} 个同行公司代码")
            elif 'code' in peer_companies.columns:
                stock_codes = peer_companies['code'].unique().tolist()[:max_stocks]
                print(f"   提取到 {len(stock_codes)} 个同行公司代码")
            else:
                print(f"   ⚠️ 同行公司数据中没有ts_code或code列")
                # 尝试查看所有列名
                print(f"   可用列名: {list(peer_companies.columns)}")
        elif industry_code:
            # 降级：使用行业成分股
            print(f"📊 使用行业指数成分股计算行业Beta：{industry_code}")
            try:
                # 获取行业成分股
                df_members = self.pro.index_member(
                    index_code=industry_code,
                    fields='con_code,index_code,in_date,out_date,is_new'
                )

                if df_members.empty:
                    return {'industry_beta': 1.0, 'error': '未找到行业成分股'}

                # 筛选当前成分股
                today = datetime.now().strftime('%Y%m%d')
                df_members = df_members[
                    (df_members['out_date'].isna()) | (df_members['out_date'] > today)
                ]

                stock_codes = df_members['con_code'].tolist()[:max_stocks]
                print(f"   行业成分股数量：{len(stock_codes)}")
            except Exception as e:
                print(f"   ⚠️ 获取行业成分股失败: {e}")
                return {'industry_beta': 1.0, 'error': str(e)}
        else:
            return {'industry_beta': 1.0, 'error': '未提供同行公司数据或行业代码'}

        if not stock_codes:
            return {'industry_beta': 1.0, 'error': '没有可用的同行公司'}

        try:
            # 计算每只股票的Beta
            betas = []
            for stock_code in stock_codes:
                result = self.calculate_beta_from_api(stock_code, window=window)  # 明确传入500天窗口
                if 'beta' in result and result['beta'] > 0:
                    betas.append(result['beta'])
                    print(f"   {stock_code}: Beta = {result['beta']:.3f}")
                else:
                    print(f"   {stock_code}: Beta计算失败或数据不足")

            if len(betas) == 0:
                return {'industry_beta': 1.0, 'error': '未成功计算任何Beta'}

            # 行业Beta = 中位数
            industry_beta = np.median(betas)

            return {
                'industry_beta': round(industry_beta, 3),
                'stock_count': len(betas),
                'beta_range': [round(min(betas), 3), round(max(betas), 3)],
                'sample_stocks': stock_codes
            }

        except Exception as e:
            print(f"⚠️ 计算行业Beta失败: {e}")
            return {'industry_beta': 1.0, 'error': str(e)}

    def get_capital_structure(
        self,
        stock_code: str,
        market_data: Optional[Dict] = None
    ) -> Dict:
        """
        获取资本结构（股权市值和有息债务）

        参数:
            stock_code: 股票代码
            market_data: 市场数据（如果提供，从中获取市值）

        返回:
            {
                'market_cap': float,           # 股权市值（元）
                'interest_bearing_debt': float,# 有息债务（元）
                'total_equity': float,         # 总权益（元）
                'debt_ratio': float,           # 债务占比
                'equity_ratio': float,         # 股权占比
                'debt_breakdown': dict         # 有息债务明细
            }
        """
        # 如果提供了market_data，从中获取市值
        if market_data and 'market_cap' in market_data:
            market_cap = market_data['market_cap']
        elif self.pro is not None:
            try:
                # 从daily_basic获取市值
                latest_date = datetime.now().strftime('%Y%m%d')
                df_basic = self.pro.daily_basic(
                    ts_code=stock_code,
                    trade_date=latest_date,
                    fields='total_mv'
                )

                if not df_basic.empty:
                    # total_mv单位是万元，转换为元
                    market_cap = df_basic['total_mv'].iloc[0] * 10000
                else:
                    # 如果没有最新数据，使用默认值
                    market_cap = 10_000_000_000  # 100亿
            except Exception as e:
                print(f"   获取市值失败: {e}，使用默认值")
                market_cap = 10_000_000_000
        else:
            # 使用默认值
            market_cap = 10_000_000_000

        # 获取有息债务和权益数据
        if self.pro is not None:
            try:
                # 从balancesheet获取最新数据
                # 有息负债包括：短期借款、长期借款、应付债券、一年内到期的非流动负债
                df_bs = self.pro.balancesheet(
                    ts_code=stock_code,
                    period='20241231',  # 使用最新年报
                    fields='ts_code,end_date,st_borr,lt_borr,bond_payable,non_cur_liab_due_1y,total_hldr_eqy_inc_min_int'
                )

                if not df_bs.empty and len(df_bs) > 0:
                    # 取最新一期数据
                    latest_bs = df_bs.iloc[-1]

                    # 有息负债明细
                    st_borr = latest_bs.get('st_borr', 0) or 0  # 短期借款
                    lt_borr = latest_bs.get('lt_borr', 0) or 0  # 长期借款
                    bond_payable = latest_bs.get('bond_payable', 0) or 0  # 应付债券
                    non_cur_liab_due_1y = latest_bs.get('non_cur_liab_due_1y', 0) or 0  # 一年内到期的非流动负债

                    # 合计有息负债
                    interest_bearing_debt = st_borr + lt_borr + bond_payable + non_cur_liab_due_1y
                    total_equity = latest_bs['total_hldr_eqy_inc_min_int']

                    debt_breakdown = {
                        '短期借款': st_borr,
                        '长期借款': lt_borr,
                        '应付债券': bond_payable,
                        '一年内到期非流动负债': non_cur_liab_due_1y,
                        '有息负债合计': interest_bearing_debt
                    }

                    print(f"   成功获取资产负债表数据（{latest_bs['end_date']}）")
                    print(f"   有息负债明细：")
                    print(f"     - 短期借款: {st_borr/1e8:.2f} 亿元")
                    print(f"     - 长期借款: {lt_borr/1e8:.2f} 亿元")
                    print(f"     - 应付债券: {bond_payable/1e8:.2f} 亿元")
                    print(f"     - 一年内到期非流动负债: {non_cur_liab_due_1y/1e8:.2f} 亿元")
                    print(f"     - 有息负债合计: {interest_bearing_debt/1e8:.2f} 亿元")

                else:
                    print(f"   未获取到资产负债表数据，使用默认值")
                    # 使用默认值
                    interest_bearing_debt = market_cap * 0.3  # 假设有息债务占市值30%
                    total_equity = market_cap
                    debt_breakdown = {}

            except Exception as e:
                print(f"   获取债务数据失败: {e}，使用默认值")
                interest_bearing_debt = market_cap * 0.3
                total_equity = market_cap
                debt_breakdown = {}
        else:
            # 使用默认值
            interest_bearing_debt = market_cap * 0.3
            total_equity = market_cap
            debt_breakdown = {}

        # 计算占比
        total_value = market_cap + interest_bearing_debt
        debt_ratio = interest_bearing_debt / total_value if total_value > 0 else 0.3
        equity_ratio = market_cap / total_value if total_value > 0 else 0.7

        return {
            'market_cap': market_cap,
            'interest_bearing_debt': interest_bearing_debt,
            'total_debt': interest_bearing_debt,  # 兼容旧字段名
            'total_equity': total_equity,
            'debt_ratio': round(debt_ratio, 3),
            'equity_ratio': round(equity_ratio, 3),
            'debt_breakdown': debt_breakdown
        }

    def calculate_wacc(
        self,
        stock_code: str,
        beta: Optional[float] = None,
        market_data: Optional[Dict] = None,
        industry_code: Optional[str] = None,
        peer_companies: Optional[pd.DataFrame] = None,
        custom_params: Optional[Dict] = None,
        use_industry_beta: bool = True
    ) -> Dict:
        """
        计算WACC（加权平均资本成本）

        参数:
            stock_code: 股票代码
            beta: Beta系数（如果为None且use_industry_beta=False，则自动计算个股Beta）
            market_data: 市场数据字典（可选）
            industry_code: 申万行业代码（可选，如果提供peer_companies则不需要）
            peer_companies: 同行公司数据（从第二章获取，优先使用）
            custom_params: 自定义参数（可选，覆盖默认值）
            use_industry_beta: 是否使用行业Beta（默认True）

        返回:
            {
                'wacc': float,                    # WACC
                'cost_of_equity': float,          # 股权成本
                'cost_of_debt': float,            # 债务成本
                'beta': float,                    # 使用的Beta
                'beta_industry': float,           # 行业Beta
                'beta_stock': float,              # 个股Beta
                'capital_structure': dict,        # 资本结构
                'parameters': dict,               # 使用的参数
                'calculation_details': str        # 计算说明
            }
        """
        # 合并默认参数和自定义参数
        params = self.DEFAULT_PARAMS.copy()
        if custom_params:
            params.update(custom_params)

        print(f"\n{'='*60}")
        print(f"WACC计算：{stock_code}")
        print(f"{'='*60}")

        # 1. 计算个股Beta（用于对比）
        print(f"\n1. 计算个股Beta系数...")
        beta_stock_result = self.calculate_beta_from_api(
            stock_code,
            window=params['beta_window']
        )
        beta_stock = beta_stock_result.get('beta', 1.0)
        if 'error' in beta_stock_result:
            print(f"   {beta_stock_result['error']}")
        else:
            print(f"   个股Beta = {beta_stock:.3f}（{beta_stock_result['window']}天窗口）")

        # 2. 计算行业Beta（优先使用）
        beta_industry = 1.0
        if peer_companies is not None and not peer_companies.empty:
            print(f"\n2. 计算行业Beta系数（使用第二章同行公司）...")
            industry_result = self.calculate_industry_beta(
                peer_companies=peer_companies,
                window=params['beta_window']
            )
            beta_industry = industry_result.get('industry_beta', 1.0)
            if 'error' not in industry_result:
                print(f"   行业Beta = {beta_industry:.3f}（{industry_result['stock_count']}只同行）")
                if 'beta_range' in industry_result:
                    print(f"   Beta范围：{industry_result['beta_range'][0]:.3f} ~ {industry_result['beta_range'][1]:.3f}")
            else:
                print(f"   {industry_result['error']}")
                beta_industry = beta_stock  # 降级使用个股Beta
        elif industry_code:
            print(f"\n2. 计算行业Beta系数（使用行业指数成分股）...")
            industry_result = self.calculate_industry_beta(
                industry_code=industry_code,
                window=params['beta_window']
            )
            beta_industry = industry_result.get('industry_beta', 1.0)
            if 'error' not in industry_result:
                print(f"   行业Beta = {beta_industry:.3f}（{industry_result['stock_count']}只成分股）")
                if 'beta_range' in industry_result:
                    print(f"   Beta范围：{industry_result['beta_range'][0]:.3f} ~ {industry_result['beta_range'][1]:.3f}")
            else:
                print(f"   {industry_result['error']}")
                beta_industry = beta_stock  # 降级使用个股Beta

        # 3. 确定使用的Beta
        if use_industry_beta and industry_code:
            beta = beta_industry
            print(f"\n3. Beta系数选择：使用行业Beta = {beta:.3f}")
        else:
            beta = beta if beta is not None else beta_stock
            print(f"\n3. Beta系数选择：{'使用提供的Beta' if beta is not None else '使用个股Beta'} = {beta:.3f}")

        # 4. 获取资本结构
        print(f"\n4. 获取资本结构...")
        capital_structure = self.get_capital_structure(stock_code, market_data)

        # 5. CAPM计算股权成本
        print(f"\n5. CAPM模型计算股权成本...")
        cost_of_equity = params['risk_free_rate'] + beta * params['market_premium']
        print(f"   无风险利率（Rf）：{params['risk_free_rate']*100:.2f}%（十年期国债收益率）")
        print(f"   市场收益率（Rm）：{params.get('market_return', (params['risk_free_rate'] + params['market_premium']))*100:.1f}%（沪深300预期收益率）")
        print(f"   市场风险溢价（Rm - Rf）：{params['market_premium']*100:.2f}%")
        print(f"   Beta系数（β）：{beta:.3f}（{'行业Beta' if use_industry_beta else '个股Beta'}）")
        print(f"   股权成本（Re）= Rf + β × (Rm - Rf)")
        print(f"                = {params['risk_free_rate']*100:.2f}% + {beta:.3f} × {params['market_premium']*100:.2f}%")
        print(f"                = {cost_of_equity*100:.2f}%")

        # 6. 计算债务成本（税前）
        print(f"\n6. 计算债务成本...")
        cost_of_debt = params['risk_free_rate'] * (1 + params['debt_premium'])
        print(f"   无风险利率：{params['risk_free_rate']*100:.2f}%")
        print(f"   债务成本上浮比例：{params['debt_premium']*100:.0f}%")
        print(f"   税前债务成本 = {params['risk_free_rate']*100:.2f}% × (1 + {params['debt_premium']*100:.0f}%)")
        print(f"                = {cost_of_debt*100:.2f}%")

        # 7. 计算WACC
        print(f"\n7. 计算WACC...")
        cost_of_debt_after_tax = cost_of_debt * (1 - params['tax_rate'])
        wacc = (
            capital_structure['equity_ratio'] * cost_of_equity +
            capital_structure['debt_ratio'] * cost_of_debt_after_tax
        )

        print(f"   企业所得税率：{params['tax_rate']:.0%}")
        print(f"   税后债务成本：{cost_of_debt_after_tax*100:.2f}%")
        print(f"   WACC = 股权占比 × 股权成本 + 债务占比 × 税后债务成本")
        print(f"       = {capital_structure['equity_ratio']:.1%} × {cost_of_equity*100:.2f}% + {capital_structure['debt_ratio']:.1%} × {cost_of_debt_after_tax*100:.2f}%")
        print(f"       = {wacc*100:.2f}%")

        # 8. 生成计算说明
        details = f"""
WACC计算详情：
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
【参数说明】
   无风险利率（Rf）：{params['risk_free_rate']*100:.2f}%（十年期国债收益率，2025年末数据）
   市场收益率（Rm）：{params.get('market_return', (params['risk_free_rate'] + params['market_premium']))*100:.1f}%（沪深300预期收益率）
   市场风险溢价（Rm - Rf）：{params['market_premium']*100:.2f}%
   企业所得税率：{params['tax_rate']:.0%}
   债务成本上浮比例：{params['debt_premium']*100:.0f}%

【Beta系数】
   个股Beta：{beta_stock:.3f}（{params['beta_window']}天窗口）
   行业Beta：{beta_industry:.3f}（行业成分股中位数）
   采用Beta：{beta:.3f}（{'行业Beta' if use_industry_beta else '个股Beta'}）

【资本结构】
   股权市值：{capital_structure['market_cap']/1e8:.2f} 亿元
   有息负债：{capital_structure['interest_bearing_debt']/1e8:.2f} 亿元"""

        if capital_structure.get('debt_breakdown'):
            details += f"""
   有息负债明细：
     - 短期借款：{capital_structure['debt_breakdown'].get('短期借款', 0)/1e8:.2f} 亿元
     - 长期借款：{capital_structure['debt_breakdown'].get('长期借款', 0)/1e8:.2f} 亿元
     - 应付债券：{capital_structure['debt_breakdown'].get('应付债券', 0)/1e8:.2f} 亿元
     - 一年内到期非流动负债：{capital_structure['debt_breakdown'].get('一年内到期非流动负债', 0)/1e8:.2f} 亿元"""

        details += f"""
   股权占比：{capital_structure['equity_ratio']*100:.1f}%
   债务占比：{capital_structure['debt_ratio']*100:.1f}%

【CAPM模型（股权成本）】
   Re = Rf + β × (Rm - Rf)
       = {params['risk_free_rate']*100:.2f}% + {beta:.3f} × {params['market_premium']*100:.2f}%
       = {cost_of_equity*100:.2f}%

【债务成本】
   税前债务成本 = Rf × (1 + 上浮比例)
                = {params['risk_free_rate']*100:.2f}% × (1 + {params['debt_premium']*100:.0f}%)
                = {cost_of_debt*100:.2f}%
   税后债务成本 = {cost_of_debt*100:.2f}% × (1 - {params['tax_rate']:.0%})
                = {cost_of_debt_after_tax*100:.2f}%

【WACC结果】
   WACC = 股权占比 × 股权成本 + 债务占比 × 税后债务成本
        = {capital_structure['equity_ratio']:.1%} × {cost_of_equity*100:.2f}% + {capital_structure['debt_ratio']:.1%} × {cost_of_debt_after_tax*100:.2f}%
        = {wacc*100:.2f}%
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

        return {
            'wacc': round(wacc, 4),
            'cost_of_equity': round(cost_of_equity, 4),
            'cost_of_debt': round(cost_of_debt, 4),
            'cost_of_debt_after_tax': round(cost_of_debt_after_tax, 4),
            'beta': beta,
            'beta_stock': beta_stock,
            'beta_industry': beta_industry,
            'capital_structure': capital_structure,
            'parameters': params,
            'calculation_details': details,
            'calculation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }


def calculate_wacc_simple(
    beta: float = 1.0,
    equity_ratio: float = 0.7,
    debt_ratio: float = 0.3,
    risk_free_rate: float = 0.0185,
    market_premium: float = 0.0615,
    debt_premium: float = 0.50,
    tax_rate: float = 0.25
) -> float:
    """
    简化的WACC计算函数（不需要API）

    参数:
        beta: Beta系数
        equity_ratio: 股权占比
        debt_ratio: 债务占比
        risk_free_rate: 无风险利率（默认1.85%，十年期国债收益率）
        market_premium: 市场风险溢价（默认6.15%）
        debt_premium: 债务成本上浮比例（默认50%）
        tax_rate: 所得税率

    返回:
        WACC（小数形式，如0.10表示10%）
    """
    cost_of_equity = risk_free_rate + beta * market_premium
    cost_of_debt = risk_free_rate * (1 + debt_premium)
    wacc = (
        equity_ratio * cost_of_equity +
        debt_ratio * cost_of_debt * (1 - tax_rate)
    )
    return wacc


if __name__ == '__main__':
    # 测试代码
    print("WACC计算器测试\n")

    # 测试简化计算
    print("1. 简化计算测试：")
    wacc_simple = calculate_wacc_simple(beta=1.2)
    print(f"   Beta=1.2时的WACC = {wacc_simple*100:.2f}%")
    print(f"   （使用新参数：Rf=1.85%, Rm=8.0%, 债务成本=Rf×1.5）\n")

    # 测试完整计算（需要tushare API）
    try:
        import tushare as ts
        pro = ts.pro_api()
        calculator = WACCCalculator(pro_api=pro)

        print("2. 完整计算测试（光弘科技 300735.SZ）：")
        result = calculator.calculate_wacc(
            '300735.SZ',
            industry_code='801010.SI',  # 电子行业
            use_industry_beta=True
        )
        print(result['calculation_details'])

    except ImportError:
        print("   未安装tushare，跳过完整计算测试")
