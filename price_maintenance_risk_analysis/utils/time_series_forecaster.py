"""
时间序列预测模块
包含ARIMA和GARCH模型用于预测漂移率和波动率
"""
import numpy as np
import pandas as pd
from typing import Dict, Optional
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

    def forecast_drift_with_arima(
        self,
        horizon: int = 120,
        order: tuple = (1, 1, 1)
    ) -> Dict:
        """
        使用ARIMA预测漂移率（年化收益率）

        Args:
            horizon: 预测期数（默认120日）
            order: ARIMA模型阶数(p,d,q)，默认(1,1,1)

        Returns:
            {
                'forecast_drift': float,      # 预测的年化漂移率
                'model_fitted': bool,         # 模型是否成功拟合
                'aic': float,                 # AIC准则（如果拟合成功）
                'forecast_series': pd.Series, # 预测序列
                'error': str                  # 错误信息（如果失败）
            }
        """
        try:
            from statsmodels.tsa.arima.model import ARIMA

            # 使用对数收益率
            log_returns = np.log(self.prices).diff().dropna()

            # 重置索引为整数索引，避免ARIMA索引警告
            log_returns = log_returns.reset_index(drop=True)

            # 拟合ARIMA模型
            model = ARIMA(log_returns, order=order)
            fitted = model.fit()

            # 预测未来horizon期
            forecast = fitted.forecast(steps=horizon)

            # 计算年化漂移率
            # 假设252个交易日/年
            total_return = forecast.sum()
            annualized_drift = total_return / horizon * 252

            return {
                'forecast_drift': annualized_drift,
                'model_fitted': True,
                'aic': fitted.aic,
                'forecast_series': forecast
            }
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
                'forecast_series': pd.Series([historical_drift] * horizon)
            }

    def forecast_volatility_with_garch(
        self,
        horizon: int = 120,
        p: int = 1,
        q: int = 1
    ) -> Dict:
        """
        使用GARCH模型预测波动率

        Args:
            horizon: 预测期数（默认120日）
            p: GARCH(p,q)的p阶数
            q: GARCH(p,q)的q阶数

        Returns:
            {
                'forecast_volatility': float,  # 预测的年化波动率
                'model_fitted': bool,          # 模型是否成功拟合
                'omega': float,                # GARCH常数项
                'alpha': float,                # ARCH系数
                'beta': float,                 # GARCH系数
                'error': str                   # 错误信息（如果失败）
            }
        """
        try:
            from arch import arch_model

            # 准备数据：转换为百分比（GARCH模型更稳定）
            returns_pct = self.prices.pct_change().dropna() * 100

            # 拟合GARCH(1,1)模型
            model = arch_model(returns_pct, vol='Garch', p=p, q=q, mean='Constant')
            fitted = model.fit(disp='off')

            # 预测未来horizon期
            forecast = fitted.forecast(horizon=horizon)
            variance_forecast = forecast.variance.iloc[-1]

            # 计算平均年化波动率
            # GARCH输出的是方差，需要开方得到波动率
            vol_daily_pct = np.sqrt(variance_forecast).mean()
            # 转换为小数形式：百分比除以100，再年化
            annualized_vol = vol_daily_pct / 100 * np.sqrt(252)

            # 提取模型参数
            params = fitted.params

            return {
                'forecast_volatility': annualized_vol,
                'model_fitted': True,
                'omega': params.get('omega', 0),
                'alpha': params.get('alpha[1]', 0),
                'beta': params.get('beta[1]', 0)
            }
        except Exception as e:
            print(f"⚠️ GARCH拟合失败: {e}")
            print(f"   使用历史波动率作为降级策略")

            # 降级策略：使用历史波动率
            historical_returns = self.prices.pct_change().dropna()
            historical_vol = historical_returns.std() * np.sqrt(252)

            return {
                'forecast_volatility': historical_vol,
                'model_fitted': False,
                'error': str(e)
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
