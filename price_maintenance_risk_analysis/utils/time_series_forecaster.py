"""
时间序列预测模块
包含ARIMA和GARCH模型用于预测漂移率和波动率
"""
import numpy as np
import pandas as pd
from typing import Dict, Optional, List, Tuple
import warnings
warnings.filterwarnings('ignore')


class TimeSeriesForecaster:
    """时间序列预测器

    使用ARIMA模型预测漂移率，GARCH模型预测波动率
    """

    def __init__(self, prices: pd.Series):
        """
        初始化预测器

        Args:
            prices: 价格序列（pd.Series）
        """
        self.prices = prices.sort_index()
        self.returns = self._calculate_returns()

    def _calculate_returns(self) -> pd.Series:
        """计算对数收益率"""
        log_prices = np.log(self.prices)
        returns = log_prices.diff().dropna()
        return returns

    def adf_test(self, series: pd.Series = None) -> Dict:
        """
        ADF检验（Augmented Dickey-Fuller Test）- 检验时间序列的平稳性

        Args:
            series: 要检验的序列，默认为对数收益率

        Returns:
            {
                'adf_statistic': float,     # ADF统计量
                'p_value': float,            # p值
                'is_stationary': bool,       # 是否平稳（p<0.05）
                'critical_values': dict,     # 临界值
                'interpretation': str        # 解释
            }
        """
        from statsmodels.tsa.stattools import adfuller

        if series is None:
            series = self._calculate_returns()

        try:
            result = adfuller(series, autolag='AIC')

            adf_stat = result[0]
            p_value = result[1]
            critical_values = result[4]

            is_stationary = p_value < 0.05

            interpretation = ""
            if is_stationary:
                interpretation = f"序列平稳（p={p_value:.4f} < 0.05），拒绝存在单位根的原假设"
            else:
                interpretation = f"序列非平稳（p={p_value:.4f} >= 0.05），不能拒绝原假设，需要差分"

            return {
                'adf_statistic': adf_stat,
                'p_value': p_value,
                'is_stationary': is_stationary,
                'critical_values': critical_values,
                'interpretation': interpretation
            }
        except Exception as e:
            print(f"⚠️ ADF检验失败: {e}")
            return {
                'error': str(e),
                'is_stationary': True  # 默认假设平稳
            }

    def find_optimal_arima_order(
        self,
        max_p: int = 3,
        max_d: int = 2,
        max_q: int = 3,
        information_criterion: str = 'aic'
    ) -> Dict:
        """
        自动搜索最优ARIMA模型阶数

        Args:
            max_p: AR阶数最大值
            max_d: 差分阶数最大值
            max_q: MA阶数最大值
            information_criterion: 信息准则 ('aic' 或 'bic')

        Returns:
            {
                'optimal_order': tuple,  # 最优阶数 (p,d,q)
                'best_ic': float,        # 最优信息准则值
                'all_results': list,     # 所有组合结果
                'top_models': list       # 前N个最优模型
            }
        """
        try:
            import itertools
            import warnings
            from statsmodels.tsa.arima.model import ARIMA

            log_returns = np.log(self.prices).diff().dropna()
            log_returns = log_returns.reset_index(drop=True)

            # 生成所有可能的(p,d,q)组合
            p_range = range(0, max_p + 1)
            d_range = range(0, max_d + 1)
            q_range = range(0, max_q + 1)

            pdq = list(itertools.product(p_range, d_range, q_range))

            results = []
            total_combinations = len(pdq)

            print(f"    正在测试 {total_combinations} 种(p,d,q)组合...")

            for i, (p, d, q) in enumerate(pdq):
                try:
                    model = ARIMA(log_returns, order=(p, d, q))

                    # 抑制收敛警告，使用默认参数
                    with warnings.catch_warnings():
                        warnings.simplefilter("ignore")
                        fitted = model.fit()

                    ic_value = fitted.aic if information_criterion == 'aic' else fitted.bic

                    results.append({
                        'order': (p, d, q),
                        'aic': fitted.aic,
                        'bic': fitted.bic,
                        information_criterion: ic_value
                    })

                    if (i + 1) % 20 == 0:
                        print(f"    进度: {i+1}/{total_combinations}")

                except Exception:
                    # 跳过无法拟合的组合
                    continue

            if not results:
                print("⚠️ 没有成功拟合的模型，使用默认(1,1,1)")
                return {
                    'optimal_order': (1, 1, 1),
                    'best_ic': float('inf'),
                    'all_results': [],
                    'top_models': []
                }

            # 按信息准则排序
            results.sort(key=lambda x: x[information_criterion])

            optimal_order = results[0]['order']
            best_ic = results[0][information_criterion]

            # 取前5个最优模型
            top_models = results[:5]

            print(f"    ✅ 最优阶数: ARIMA{optimal_order}, {information_criterion.upper()}={best_ic:.2f}")

            return {
                'optimal_order': optimal_order,
                'best_ic': best_ic,
                'all_results': results,
                'top_models': top_models
            }

        except Exception as e:
            print(f"⚠️ 自动寻优失败: {e}")
            return {
                'optimal_order': (1, 1, 1),
                'best_ic': float('inf'),
                'all_results': [],
                'top_models': [],
                'error': str(e)
            }

    def forecast_drift_with_arima(
        self,
        horizon: int = 120,
        order: tuple = None,
        auto_find_order: bool = True,
        max_p: int = 3,
        max_d: int = 2,
        max_q: int = 3
    ) -> Dict:
        """
        使用ARIMA预测漂移率（年化收益率）

        Args:
            horizon: 预测期数（默认120日）
            order: ARIMA模型阶数(p,d,q)，如果为None且auto_find_order=True则自动寻优
            auto_find_order: 是否自动寻找最优阶数
            max_p, max_d, max_q: 自动寻优时的最大阶数

        Returns:
            {
                'forecast_drift': float,      # 预测的年化漂移率
                'model_fitted': bool,         # 模型是否成功拟合
                'aic': float,                 # AIC准则（如果拟合成功）
                'forecast_series': pd.Series, # 预测序列
                'fitted_model': object,       # 拟合的模型对象
                'log_returns': pd.Series,    # 对数收益率序列
                'order_used': tuple,          # 实际使用的阶数
                'order_selection': dict,      # 阶数选择结果（如果自动寻优）
                'error': str                  # 错误信息（如果失败）
            }
        """
        try:
            from statsmodels.tsa.arima.model import ARIMA
            import warnings

            # 使用对数收益率
            log_returns = np.log(self.prices).diff().dropna()
            log_returns = log_returns.reset_index(drop=True)

            # 确定使用的阶数
            order_selection = None
            if order is None and auto_find_order:
                print("\n    自动寻找最优ARIMA阶数...")
                order_selection = self.find_optimal_arima_order(
                    max_p=max_p, max_d=max_d, max_q=max_q
                )
                order = order_selection['optimal_order']
                print(f"    使用最优阶数: ARIMA{order}")
            elif order is None:
                order = (1, 1, 1)  # 默认值

            # 拟合ARIMA模型
            model = ARIMA(log_returns, order=order)

            # 抑制收敛警告，使用默认参数
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                fitted = model.fit()

            # 预测未来horizon期
            forecast = fitted.forecast(steps=horizon)

            # 计算年化漂移率（两种方式）
            # 方式1：简单年化（对数收益率线性年化）
            total_log_return = forecast.sum()
            annualized_drift_simple = total_log_return / horizon * 252

            # 方式2：复利年化（更准确，转换为简单收益率后年化再转回）
            # 累积简单收益率
            total_simple_return = np.exp(total_log_return) - 1
            # 复利年化
            annualized_simple_return = (1 + total_simple_return) ** (252 / horizon) - 1
            # 转回对数收益率
            annualized_drift_compound = np.log(1 + annualized_simple_return)

            # 使用复利年化方式（更准确）
            annualized_drift = annualized_drift_compound

            result = {
                'forecast_drift': annualized_drift,
                'forecast_drift_simple': annualized_drift_simple,  # 保留简单年化供参考
                'forecast_drift_compound': annualized_drift_compound,  # 复利年化
                'total_log_return': total_log_return,
                'total_simple_return': total_simple_return,
                'model_fitted': True,
                'aic': fitted.aic,
                'forecast_series': forecast,
                'fitted_model': fitted,
                'log_returns': log_returns,
                'order_used': order
            }

            if order_selection:
                result['order_selection'] = order_selection

            return result

        except Exception as e:
            print(f"⚠️ ARIMA拟合失败: {e}")
            print(f"   使用历史平均收益率作为降级策略")

            # 降级策略：使用历史平均收益率
            historical_returns = self.prices.pct_change().dropna()
            historical_drift = historical_returns.mean() * 252

            return {
                'forecast_drift': historical_drift,
                'model_fitted': False,
                'error': str(e),
                'forecast_series': pd.Series([historical_drift] * horizon),
                'order_used': order if order else (1, 1, 1)
            }

    def arch_lm_test(self, series: pd.Series = None) -> Dict:
        """
        ARCH-LM检验 - 检验时间序列是否存在ARCH效应

        Args:
            series: 要检验的序列，默认为收益率序列

        Returns:
            {
                'lm_statistic': float,      # LM统计量
                'p_value': float,            # p值
                'has_arch_effect': bool,     # 是否存在ARCH效应
                'interpretation': str        # 解释
            }
        """
        try:
            from statsmodels.stats.diagnostic import het_arch

            if series is None:
                series = self.prices.pct_change().dropna() * 100

            # 执行ARCH-LM检验
            lm_stat, p_value, _, _ = het_arch(series, nlags=10)

            has_arch_effect = p_value < 0.05

            interpretation = ""
            if has_arch_effect:
                interpretation = f"存在ARCH效应（p={p_value:.4f} < 0.05），拒绝原假设，适合使用GARCH模型"
            else:
                interpretation = f"不存在ARCH效应（p={p_value:.4f} >= 0.05），无需使用GARCH模型"

            return {
                'lm_statistic': lm_stat,
                'p_value': p_value,
                'has_arch_effect': has_arch_effect,
                'interpretation': interpretation
            }
        except Exception as e:
            print(f"⚠️ ARCH-LM检验失败: {e}")
            return {
                'error': str(e),
                'has_arch_effect': True  # 默认假设有ARCH效应
            }

    def find_optimal_garch_order(
        self,
        max_p: int = 2,
        max_q: int = 2,
        information_criterion: str = 'aic'
    ) -> Dict:
        """
        自动搜索最优GARCH模型阶数

        Args:
            max_p: GARCH的p阶数最大值
            max_q: GARCH的q阶数最大值
            information_criterion: 信息准则 ('aic' 或 'bic')

        Returns:
            {
                'optimal_order': tuple,  # 最优阶数 (p,q)
                'best_ic': float,        # 最优信息准则值
                'all_results': list,     # 所有组合结果
                'top_models': list       # 前N个最优模型
            }
        """
        try:
            import itertools
            from arch import arch_model

            returns_pct = self.prices.pct_change().dropna() * 100

            # 生成所有可能的(p,q)组合
            p_range = range(1, max_p + 1)
            q_range = range(1, max_q + 1)

            pq_combinations = list(itertools.product(p_range, q_range))

            results = []
            total_combinations = len(pq_combinations)

            print(f"    正在测试 {total_combinations} 种(p,q)组合...")

            for i, (p, q) in enumerate(pq_combinations):
                try:
                    model = arch_model(returns_pct, vol='Garch', p=p, q=q, mean='Constant')
                    fitted = model.fit(disp='off')

                    ic_value = fitted.aic if information_criterion == 'aic' else fitted.bic

                    # 提取参数
                    params = fitted.params
                    omega = params.get('omega', 0)
                    alpha = params.get('alpha[1]', 0) if p >= 1 else 0
                    beta = params.get('beta[1]', 0) if q >= 1 else 0
                    persistence = alpha + beta

                    results.append({
                        'order': (p, q),
                        'aic': fitted.aic,
                        'bic': fitted.bic,
                        information_criterion: ic_value,
                        'omega': omega,
                        'alpha': alpha,
                        'beta': beta,
                        'persistence': persistence
                    })

                    if (i + 1) % 5 == 0:
                        print(f"    进度: {i+1}/{total_combinations}")

                except Exception:
                    # 跳过无法拟合的组合
                    continue

            if not results:
                print("⚠️ 没有成功拟合的模型，使用默认(1,1)")
                return {
                    'optimal_order': (1, 1),
                    'best_ic': float('inf'),
                    'all_results': [],
                    'top_models': []
                }

            # 按信息准则排序
            results.sort(key=lambda x: x[information_criterion])

            optimal_order = results[0]['order']
            best_ic = results[0][information_criterion]

            # 取前5个最优模型
            top_models = results[:5]

            print(f"    ✅ 最优阶数: GARCH({optimal_order[0]},{optimal_order[1]}), {information_criterion.upper()}={best_ic:.2f}")

            return {
                'optimal_order': optimal_order,
                'best_ic': best_ic,
                'all_results': results,
                'top_models': top_models
            }

        except Exception as e:
            print(f"⚠️ 自动寻优失败: {e}")
            return {
                'optimal_order': (1, 1),
                'best_ic': float('inf'),
                'all_results': [],
                'top_models': [],
                'error': str(e)
            }

    def forecast_volatility_with_garch(
        self,
        horizon: int = 120,
        p: int = None,
        q: int = None,
        auto_find_order: bool = True,
        max_p: int = 2,
        max_q: int = 2
    ) -> Dict:
        """
        使用GARCH模型预测波动率

        Args:
            horizon: 预测期数（默认120日）
            p: GARCH(p,q)的p阶数
            q: GARCH(p,q)的q阶数
            auto_find_order: 是否自动寻找最优阶数
            max_p, max_q: 自动寻优时的最大阶数

        Returns:
            {
                'forecast_volatility': float,  # 预测的年化波动率
                'model_fitted': bool,          # 模型是否成功拟合
                'omega': float,                # GARCH常数项
                'alpha': float,                # ARCH系数
                'beta': float,                 # GARCH系数
                'persistence': float,          # α+β持续性
                'conditional_variances': pd.Series,  # 条件方差序列
                'standardized_residuals': pd.Series,  # 标准化残差
                'order_used': tuple,           # 实际使用的阶数
                'order_selection': dict,       # 阶数选择结果（如果自动寻优）
                'fitted_model': object,        # 拟合的模型对象
                'error': str                   # 错误信息（如果失败）
            }
        """
        try:
            from arch import arch_model

            returns_pct = self.prices.pct_change().dropna() * 100

            # 确定使用的阶数
            order_selection = None
            if p is None or q is None:
                if auto_find_order:
                    print("\n    自动寻找最优GARCH阶数...")
                    order_selection = self.find_optimal_garch_order(
                        max_p=max_p, max_q=max_q
                    )
                    p, q = order_selection['optimal_order']
                    print(f"    使用最优阶数: GARCH({p},{q})")
                else:
                    p, q = 1, 1

            # 拟合GARCH(p,q)模型
            model = arch_model(returns_pct, vol='Garch', p=p, q=q, mean='Constant')
            fitted = model.fit(disp='off')

            # 提取模型参数
            params = fitted.params
            omega = params.get('omega', 0)
            alpha = params.get('alpha[1]', 0) if p >= 1 else 0
            beta = params.get('beta[1]', 0) if q >= 1 else 0
            persistence = alpha + beta

            # 预测未来horizon期
            forecast = fitted.forecast(horizon=horizon)
            variance_forecast = forecast.variance.iloc[-1]

            # 计算平均年化波动率
            vol_daily_pct = np.sqrt(variance_forecast).mean()
            annualized_vol = vol_daily_pct / 100 * np.sqrt(252)

            # 获取条件方差和标准化残差
            conditional_variances = fitted.conditional_volatility ** 2
            standardized_residuals = fitted.resid / fitted.conditional_volatility

            result = {
                'forecast_volatility': annualized_vol,
                'model_fitted': True,
                'omega': omega,
                'alpha': alpha,
                'beta': beta,
                'persistence': persistence,
                'conditional_variances': conditional_variances,
                'standardized_residuals': standardized_residuals,
                'order_used': (p, q),
                'fitted_model': fitted
            }

            if order_selection:
                result['order_selection'] = order_selection

            return result

        except Exception as e:
            print(f"⚠️ GARCH拟合失败: {e}")
            print(f"   使用历史波动率作为降级策略")

            # 降级策略：使用历史波动率
            historical_returns = self.prices.pct_change().dropna()
            historical_vol = historical_returns.std() * np.sqrt(252)

            return {
                'forecast_volatility': historical_vol,
                'model_fitted': False,
                'error': str(e),
                'order_used': (p if p else 1, q if q else 1)
            }

    def forecast_both(
        self,
        horizon: int = 120
    ) -> Dict:
        """
        同时预测漂移率和波动率

        Args:
            horizon: 预测期数

        Returns:
            {
                'drift': dict,  # ARIMA预测结果
                'volatility': dict  # GARCH预测结果
            }
        """
        print(f"\n{'='*60}")
        print(f"📊 时间序列预测（预测期数：{horizon}日）")
        print(f"{'='*60}")

        # 预测漂移率
        print("\n1️⃣ ARIMA模型预测漂移率...")
        drift_result = self.forecast_drift_with_arima(horizon=horizon)
        status = "✅ 成功" if drift_result['model_fitted'] else "⚠️ 降级"
        print(f"   状态: {status}")
        print(f"   预测漂移率: {drift_result['forecast_drift']*100:.2f}%（年化）")
        if drift_result['model_fitted']:
            print(f"   AIC: {drift_result['aic']:.2f}")

        # 预测波动率
        print("\n2️⃣ GARCH模型预测波动率...")
        vol_result = self.forecast_volatility_with_garch(horizon=horizon)
        status = "✅ 成功" if vol_result['model_fitted'] else "⚠️ 降级"
        print(f"   状态: {status}")
        print(f"   预测波动率: {vol_result['forecast_volatility']*100:.2f}%（年化）")
        if vol_result['model_fitted']:
            print(f"   模型参数: ω={vol_result['omega']:.4f}, "
                  f"α={vol_result['alpha']:.4f}, β={vol_result['beta']:.4f}")
            print(f"   持续性 α+β={vol_result['alpha'] + vol_result['beta']:.4f}")

        print(f"\n{'='*60}")

        return {
            'drift': drift_result,
            'volatility': vol_result
        }


# 便捷函数
def create_forecaster_from_market_data(market_data: Dict) -> Optional[TimeSeriesForecaster]:
    """
    从market_data字典创建预测器

    Args:
        market_data: 包含price_series的字典

    Returns:
        TimeSeriesForecaster实例，如果数据不足则返回None
    """
    prices = market_data.get('price_series', [])

    if not prices or len(prices) < 100:
        print("⚠️ 价格序列数据不足（需要至少100个数据点）")
        return None

    prices_series = pd.Series(prices)

    return TimeSeriesForecaster(prices_series)
