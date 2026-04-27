"""
定增项目风险分析工具函数
包含各种风险分析的核心计算函数

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【重要说明】收益率类型与GBM模型

本模块中的GBM模型相关函数使用**连续复利（对数收益率）**。

报告中的收益率数据是**离散复利**（简单涨跌幅），用于直观展示。
在GBM模型中使用时需要转换为连续复利。

【转换方法】
    from utils.return_conversion import discrete_to_continuous, get_gbm_drift_from_discrete

    # 场景1: 从报告数据（离散复利）转换为GBM漂移率（连续复利）
    period_return_discrete = market_data['period_return_60d']  # 例如 -0.1462
    drift = get_gbm_drift_from_discrete(period_return_discrete, 60)
    # = ln(1 - 0.1462) × (250/60) = -0.1580 × 4.167 = -0.658

    # 场景2: 蒙特卡洛模拟使用
    drift = get_gbm_drift_from_discrete(period_return, window)
    volatility = market_data['volatility_60d']  # 已是连续复利
    monte_carlo_simulation(drift=drift, volatility=volatility)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
import numpy as np
import pandas as pd
from scipy import stats
from typing import Dict, List, Tuple, Optional
import warnings
warnings.filterwarnings('ignore')


def calculate_profit_probability_lognormal(target_price: float,
                                           current_price: float,
                                           drift: float,
                                           volatility: float,
                                           period_months: int) -> float:
    """
    使用对数正态分布计算盈利概率（几何布朗运动模型）

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    【重要说明】参数类型要求

    本函数基于GBM模型，要求输入**连续复利**的漂移率和波动率。

    参数要求：
        - drift: 连续复利年化漂移率 μ（不是离散复利！）
        - volatility: 连续复利年化波动率 σ

    【从报告数据转换】
    如果从市场数据报告（离散复利）获取参数，需要转换：

        from utils.return_conversion import get_gbm_drift_from_discrete

        # 错误示例：直接使用离散复利
        drift_wrong = market_data['annual_return_60d']  # ❌ 这是离散复利

        # 正确示例：转换为连续复利
        period_return = market_data['period_return_60d']  # -14.62%
        drift_correct = get_gbm_drift_from_discrete(period_return, 60)
        # = ln(1 - 0.1462) × (250/60) = -60.91%（连续复利年化）

        volatility = market_data['volatility_60d']  # ✓ 已是连续复利

    ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    理论基础：
    - 股票价格服从几何布朗运动：dS = μSdt + σSdW
    - 对数价格服从正态分布：ln(S_t) ~ N(ln(S_0) + (μ - σ²/2)t, σ²t)
    - 盈利概率：P(S_t > target_price)

    参数:
        target_price: 目标价格（盈利阈值）
        current_price: 当前价格
        drift: 年化漂移率 μ（必须是连续复利！）
        volatility: 年化波动率 σ（连续复利）
        period_months: 预测期（月）

    返回:
        float: 盈利概率（0-100）

    示例:
        >>> from utils.return_conversion import get_gbm_drift_from_discrete
        >>> period_return = -0.1462  # 60日离散收益率 -14.62%
        >>> drift = get_gbm_drift_from_discrete(period_return, 60)
        >>> volatility = 0.338  # 60日波动率 33.80%
        >>>
        >>> prob = calculate_profit_probability_lognormal(
        ...     target_price=30.0,
        ...     current_price=21.26,
        ...     drift=drift,
        ...     volatility=volatility,
        ...     period_months=6
        ... )
        >>> print(f"盈利概率: {prob:.2f}%")
    """
    T = period_months / 12  # 转换为年

    # 对数正态分布参数
    # E[ln(S_t)] = ln(S_0) + (μ - σ²/2)t
    log_mean = np.log(current_price) + (drift - volatility**2 / 2) * T
    log_std = volatility * np.sqrt(T)

    # P(S_t > target_price) = P(ln(S_t) > ln(target_price))
    z_score = (np.log(target_price) - log_mean) / log_std
    probability = (1 - stats.norm.cdf(z_score)) * 100

    return probability


class PrivatePlacementRiskAnalyzer:
    """定增项目风险分析器"""

    def __init__(self,
                 issue_price: float,
                 issue_shares: int,
                 lockup_period: int,  # 锁定期（月）
                 current_price: float = None,
                 risk_free_rate: float = 0.03):
        """
        初始化定增项目参数

        Args:
            issue_price: 发行价格（元/股）
            issue_shares: 发行数量（股）
            lockup_period: 锁定期（月）
            current_price: 当前价格（元/股）
            risk_free_rate: 无风险利率
        """
        self.issue_price = issue_price
        self.issue_shares = issue_shares
        self.lockup_period = lockup_period
        self.current_price = current_price or issue_price
        self.risk_free_rate = risk_free_rate
        self.issue_amount = issue_price * issue_shares  # 融资金额

    def calculate_break_even_price(self, discount_rate: float = 0.2) -> float:
        """
        计算盈亏平衡价格

        Args:
            discount_rate: 期望年化收益率（如20%输入0.2）

        Returns:
            盈亏平衡价格
        """
        monthly_return = (1 + discount_rate) ** (1/12) - 1
        break_even = self.issue_price * (1 + monthly_return) ** self.lockup_period
        return break_even

    def calculate_probability_above_price(self,
                                        current_price: float,
                                        volatility: float,
                                        target_price: float,
                                        period_months: int) -> float:
        """
        计算股价达到目标价位的概率

        Args:
            current_price: 当前价格
            volatility: 年化波动率
            target_price: 目标价格
            period_months: 预测期（月）

        Returns:
            概率值（0-1）
        """
        T = period_months / 12
        drift = self.risk_free_rate - 0.5 * volatility ** 2

        # 使用对数正态分布计算概率
        log_target = np.log(target_price / current_price)
        mean = np.log(current_price) + (drift + 0.5 * volatility ** 2) * T
        std = volatility * np.sqrt(T)

        probability = 1 - stats.lognorm.cdf(target_price, s=std, scale=np.exp(mean - std**2/2))
        return probability

    def monte_carlo_simulation(self,
                             n_simulations: int = 10000,
                             time_steps: int = 252,
                             volatility: float = 0.3,
                             drift: float = None,
                             seed: int = None) -> pd.DataFrame:
        """
        蒙特卡洛模拟股价路径（基于GBM模型）

        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

        【重要说明】参数类型要求

        本方法基于GBM模型，要求输入**连续复利**的漂移率和波动率。

        参数要求：
            - drift: 连续复利年化漂移率 μ（不是离散复利！）
            - volatility: 连续复利年化波动率 σ

        【从报告数据转换】
        如果从市场数据报告（离散复利）获取参数，需要转换：

            from utils.return_conversion import get_gbm_drift_from_discrete

            # 错误示例：直接使用离散复利
            drift_wrong = market_data['annual_return_60d']  # ❌ 这是离散复利

            # 正确示例：转换为连续复利
            period_return = market_data['period_return_60d']  # -14.62%
            drift_correct = get_gbm_drift_from_discrete(period_return, 60)
            # = ln(1 - 0.1462) × (250/60) = -60.91%（连续复利年化）

            volatility = market_data['volatility_60d']  # ✓ 已是连续复利

        ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

        Args:
            n_simulations: 模拟次数（路径数量）
            time_steps: 时间步数（天），默认252（一年）
            volatility: 年化波动率 σ（连续复利）
            drift: 漂移率 μ（年化，连续复利），默认使用无风险利率
            seed: 随机种子（可重复性）

        Returns:
            DataFrame: 模拟结果，每一行是一条路径，列是 Day_0, Day_1, ...

        示例:
            >>> from utils.return_conversion import get_gbm_drift_from_discrete
            >>>
            >>> # 从市场数据获取参数（离散复利）
            >>> period_return = market_data['period_return_60d']  # -14.62%
            >>> volatility = market_data['volatility_60d']  # 33.80%
            >>>
            >>> # 转换为连续复利（用于GBM）
            >>> drift = get_gbm_drift_from_discrete(period_return, 60)
            >>>
            >>> # 运行蒙特卡洛模拟
            >>> analyzer = PrivatePlacementRiskAnalyzer(...)
            >>> simulations = analyzer.monte_carlo_simulation(
            ...     n_simulations=10000,
            ...     time_steps=180,  # 6个月约180天
            ...     drift=drift,  # 连续复利年化漂移率
            ...     volatility=volatility,  # 连续复利年化波动率
            ...     seed=42  # 可重复
            ... )

        理论基础：
            GBM模型离散化形式：
                S_{t+1} = S_t × exp((μ - σ²/2)dt + σ√dt × Z)

            其中：
                - μ: 连续复利漂移率（年化）
                - σ: 连续复利波动率（年化）
                - dt: 时间步长（1/252，日数据）
                - Z: 标准正态分布随机变量
        """
        if seed is not None:
            np.random.seed(seed)

        if drift is None:
            drift = self.risk_free_rate

        dt = 1 / 252  # 日时间步长（252个交易日/年）

        # 使用传入的time_steps，如果没有则使用锁定期天数
        if time_steps is not None:
            total_days = time_steps
        else:
            total_days = self.lockup_period * 30  # 锁定期天数

        # 生成随机路径
        brownian_motion = np.random.standard_normal((n_simulations, total_days))

        # 股价路径模拟（几何布朗运动）
        # 注意：drift和volatility都应该是连续复利年化值
        prices = np.zeros((n_simulations, total_days + 1))
        prices[:, 0] = self.current_price

        for t in range(1, total_days + 1):
            prices[:, t] = prices[:, t-1] * np.exp(
                (drift - 0.5 * volatility**2) * dt +
                volatility * np.sqrt(dt) * brownian_motion[:, t-1]
            )

        # 转换为DataFrame
        columns = [f'Day_{i}' for i in range(total_days + 1)]
        df = pd.DataFrame(prices, columns=columns)
        return df

    def calculate_var(self,
                     returns: pd.Series,
                     confidence_level: float = 0.95,
                     method: str = 'historical') -> Dict[str, float]:
        """
        计算风险价值 (VaR)

        Args:
            returns: 收益率序列
            confidence_level: 置信水平（默认95%）
            method: 计算方法 ('historical', 'parametric', 'monte_carlo')

        Returns:
            VaR计算结果字典
        """
        alpha = 1 - confidence_level

        if method == 'historical':
            var = returns.quantile(alpha)
            cvar = returns[returns <= var].mean()

        elif method == 'parametric':
            # 假设收益率服从正态分布
            mu = returns.mean()
            sigma = returns.std()
            var = mu + sigma * stats.norm.ppf(alpha)
            cvar = mu - sigma * (stats.norm.pdf(stats.norm.ppf(alpha)) / alpha)

        elif method == 'monte_carlo':
            n_sims = 10000
            mu = returns.mean()
            sigma = returns.std()
            simulated = np.random.normal(mu, sigma, n_sims)
            var = np.percentile(simulated, alpha * 100)
            cvar = simulated[simulated <= var].mean()

        else:
            raise ValueError(f"Unknown method: {method}")

        return {
            'VaR': var,
            'CVaR': cvar,
            'confidence_level': confidence_level,
            'method': method
        }

    def calculate_dcf_value(self,
                           free_cash_flows: List[float],
                           terminal_growth_rate: float = 0.025,
                           wacc: float = 0.10) -> Dict[str, float]:
        """
        DCF估值（定增项目专用）

        Args:
            free_cash_flows: 未来现金流预测（万元）
            terminal_growth_rate: 终值增长率
            wacc: 加权平均资本成本

        Returns:
            估值结果
        """
        # 预测期现值
        pv_forecasts = sum([fcf / ((1 + wacc) ** (i + 1))
                          for i, fcf in enumerate(free_cash_flows)])

        # 终值
        terminal_fcf = free_cash_flows[-1] * (1 + terminal_growth_rate)
        terminal_value = terminal_fcf / (wacc - terminal_growth_rate)
        pv_terminal = terminal_value / ((1 + wacc) ** len(free_cash_flows))

        # 企业价值
        enterprise_value = pv_forecasts + pv_terminal

        return {
            'enterprise_value': enterprise_value,
            'pv_forecasts': pv_forecasts,
            'pv_terminal': pv_terminal,
            'terminal_value': terminal_value
        }

    def sensitivity_analysis(self,
                           params: Dict[str, Tuple[float, float]],
                           base_value: float,
                           n_steps: int = 100) -> pd.DataFrame:
        """
        敏感性分析

        Args:
            params: 参数字典，格式为 {'param_name': (min, max)}
            base_value: 基准值
            n_steps: 每个参数的步数

        Returns:
            敏感性分析结果
        """
        results = {}

        for param_name, (min_val, max_val) in params.items():
            values = np.linspace(min_val, max_val, n_steps)
            # 这里需要根据具体模型计算敏感性
            # 简化示例：使用线性关系
            results[param_name] = values

        return pd.DataFrame(results)

    def stress_test(self,
                   scenarios: Dict[str, Dict[str, float]]) -> pd.DataFrame:
        """
        压力测试

        Args:
            scenarios: 压力测试情景字典
                    {
                        '情景名称': {
                            'price_drop': 股价下跌幅度,
                            'volatility_spike': 波动率飙升幅度,
                            ...
                        }
                    }

        Returns:
            压力测试结果
        """
        results = []

        for scenario_name, params in scenarios.items():
            # 计算压力情景下的盈亏
            stressed_price = self.current_price * (1 - params.get('price_drop', 0))
            pnl = (stressed_price - self.issue_price) * self.issue_shares
            pnl_pct = (stressed_price / self.issue_price - 1) * 100

            results.append({
                'scenario': scenario_name,
                'stressed_price': stressed_price,
                'pnl': pnl,
                'pnl_percent': pnl_pct,
                'break_even_distance': (self.calculate_break_even_price() - stressed_price) / stressed_price
            })

        return pd.DataFrame(results)

    def calculate_max_drawdown(self, price_series: pd.Series) -> Dict[str, float]:
        """
        计算最大回撤

        Args:
            price_series: 价格序列

        Returns:
            最大回撤指标
        """
        cumulative = (1 + price_series.pct_change()).cumprod()
        running_max = cumulative.expanding().max()
        drawdown = (cumulative - running_max) / running_max

        max_dd = drawdown.min()
        max_dd_date = drawdown.idxmin()

        return {
            'max_drawdown': max_dd,
            'max_drawdown_date': max_dd_date,
            'max_drawdown_pct': max_dd * 100
        }

    def calculate_irr(self,
                      cash_flows: List[float],
                      investment: float) -> float:
        """
        计算内部收益率 (IRR)

        Args:
            cash_flows: 现金流序列（负数表示流出，正数表示流入）
            investment: 初始投资金额

        Returns:
            IRR（年化）
        """
        all_flows = [-investment] + cash_flows
        irr = np.irr(all_flows)
        return irr

    def calculate_sharpe_ratio(self,
                             returns: pd.Series,
                             risk_free_rate: float = None) -> float:
        """
        计算夏普比率

        Args:
            returns: 收益率序列
            risk_free_rate: 无风险利率（年化）

        Returns:
            夏普比率
        """
        if risk_free_rate is None:
            risk_free_rate = self.risk_free_rate

        excess_returns = returns - risk_free_rate / 250
        sharpe = excess_returns.mean() / excess_returns.std() * np.sqrt(250)
        return sharpe

    def generate_summary_report(self) -> Dict:
        """
        生成项目摘要报告

        Returns:
            摘要信息字典
        """
        # 计算关键指标
        discount_to_current = (self.current_price - self.issue_price) / self.issue_price
        break_even_20 = self.calculate_break_even_price(0.20)
        distance_to_be = (self.current_price - break_even_20) / self.current_price

        return {
            'project_info': {
                'issue_price': self.issue_price,
                'current_price': self.current_price,
                'issue_shares': self.issue_shares,
                'issue_amount': self.issue_amount,
                'lockup_period_months': self.lockup_period
            },
            'key_metrics': {
                'discount_to_current': f"{discount_to_current*100:.2f}%",
                'break_even_price_20pct': break_even_20,
                'distance_to_break_even': f"{distance_to_be*100:.2f}%"
            },
            'risk_indicators': {
                'lockup_risk': 'High' if self.lockup_period > 12 else 'Medium',
                'price_volatility_risk': 'To be calculated'
            }
        }


# 辅助函数
def load_stock_data(file_path: str) -> pd.DataFrame:
    """
    加载股票数据

    Args:
        file_path: CSV文件路径

    Returns:
        股票数据DataFrame
    """
    df = pd.read_csv(file_path)
    df['date'] = pd.to_datetime(df['date'])
    df.set_index('date', inplace=True)
    return df


def calculate_volatility(returns: pd.Series, annualize: bool = True) -> float:
    """
    计算波动率

    Args:
        returns: 收益率序列
        annualize: 是否年化

    Returns:
        波动率
    """
    vol = returns.std()
    if annualize:
        vol = vol * np.sqrt(250)
    return vol


def format_number(num: float, precision: int = 2) -> str:
    """格式化数字显示"""
    return f"{num:.{precision}f}"
