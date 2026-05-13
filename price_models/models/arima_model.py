"""ARIMA模型封装 — 线性分量预测"""

import numpy as np
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

STATSMODELS_AVAILABLE = False
try:
    from statsmodels.tsa.arima.model import ARIMA
    from statsmodels.tsa.stattools import adfuller
    STATSMODELS_AVAILABLE = True
except ImportError:
    pass


def find_optimal_order(log_returns, max_p=3, max_q=3, criterion='aic'):
    """自动搜索最优ARIMA(p,0,q)阶数（d固定为0，输入已是log returns）

    Returns:
        dict: {'order': tuple, 'aic': float, 'all_results': list}
    """
    if not STATSMODELS_AVAILABLE:
        return {'order': (1, 0, 1), 'aic': float('inf'), 'all_results': []}

    import itertools
    log_returns = np.array(log_returns, dtype=np.float64)
    results = []

    for p, q in itertools.product(range(max_p + 1), range(max_q + 1)):
        if p == 0 and q == 0:
            continue
        try:
            model = ARIMA(log_returns, order=(p, 0, q))
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                fitted = model.fit()
            ic = fitted.aic if criterion == 'aic' else fitted.bic
            results.append({'order': (p, 0, q), 'aic': fitted.aic, 'bic': fitted.bic})
        except Exception:
            continue

    if not results:
        return {'order': (1, 0, 1), 'aic': float('inf'), 'all_results': []}

    results.sort(key=lambda x: x[criterion])
    return {
        'order': results[0]['order'],
        'aic': results[0][criterion],
        'all_results': results,
    }


def fit_arima(log_returns, order=None, auto_find=True, max_p=3, max_q=3):
    """拟合ARIMA模型并返回拟合结果

    Args:
        log_returns: 对数收益率序列
        order: (p,0,q)，None则自动寻优
        auto_find: 是否自动寻优

    Returns:
        dict: {
            'fitted': model object,
            'residuals': np.array,
            'fitted_values': np.array,
            'order': tuple,
            'aic': float,
        }
    """
    if not STATSMODELS_AVAILABLE:
        raise ImportError("statsmodels not installed")

    lr = np.array(log_returns, dtype=np.float64)

    if order is None and auto_find:
        sel = find_optimal_order(lr, max_p=max_p, max_q=max_q)
        order = sel['order']
    elif order is None:
        order = (1, 0, 1)

    model = ARIMA(lr, order=order)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        fitted = model.fit()

    return {
        'fitted': fitted,
        'residuals': fitted.resid,
        'fitted_values': fitted.fittedvalues,
        'order': order,
        'aic': fitted.aic,
    }


def forecast_arima(fitted_model, horizon=120):
    """ARIMA多步预测

    Returns:
        dict: {
            'forecast': np.array (horizon,),
            'annualized_drift': float,
        }
    """
    forecast = fitted_model.forecast(steps=horizon)
    forecast = np.array(forecast)

    # 复利年化
    total_log_return = forecast.sum()
    total_simple = np.exp(total_log_return) - 1
    annualized_simple = (1 + total_simple) ** (252 / horizon) - 1
    annualized_drift = np.log(1 + annualized_simple)

    return {
        'forecast': forecast,
        'annualized_drift': annualized_drift,
        'total_log_return': total_log_return,
    }
