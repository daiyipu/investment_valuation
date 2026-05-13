"""ARIMA-LSTM 混合模型 — 残差修正方案

经典 Zhang (2003) 方案：
1. ARIMA拟合log returns → 线性分量 + 残差
2. LSTM学习残差中的非线性模式
3. Combined = ARIMA_forecast + blend_weight * LSTM_residual_forecast
"""

import numpy as np
import pandas as pd
from models.lstm_net import (TORCH_AVAILABLE, LSTMResidualNet, prepare_sequences,
                              create_dataloaders, standardize, train_lstm,
                              forecast_autoregressive)
from models.arima_model import fit_arima, forecast_arima, STATSMODELS_AVAILABLE
from config import (SEQ_LENGTH, HORIZON, LSTM_HIDDEN, LSTM_LAYERS, LSTM_DROPOUT,
                    EPOCHS, BATCH_SIZE, LEARNING_RATE, WEIGHT_DECAY, GRADIENT_CLIP,
                    EARLY_STOP_PATIENCE, LR_PATIENCE, LR_FACTOR, MIN_LR,
                    TRAIN_RATIO, FORECAST_CLIP_STD,
                    OVERFIT_THRESHOLD_HIGH, OVERFIT_THRESHOLD_LOW)


def forecast_hybrid(log_returns, horizon=HORIZON, seq_length=SEQ_LENGTH,
                    arima_order=None, auto_find_arima=True,
                    hidden_dim=LSTM_HIDDEN, num_layers=LSTM_LAYERS,
                    dropout=LSTM_DROPOUT, epochs=EPOCHS,
                    blend_weight=None, verbose=True):
    """ARIMA-LSTM混合模型预测漂移率

    Args:
        log_returns: 对数收益率序列
        horizon: 预测步数
        seq_length: LSTM输入序列长度
        arima_order: ARIMA(p,0,q)阶数，None则自动寻优
        auto_find_arima: 是否自动寻优ARIMA阶数
        blend_weight: LSTM修正权重，None则自适应
        verbose: 是否打印过程

    Returns:
        dict: {
            'forecast': np.array (horizon,),
            'annualized_drift': float,        # 混合年化漂移率
            'arima_drift': float,             # 纯ARIMA年化漂移率
            'lstm_correction': float,         # LSTM年化修正量
            'blend_weight': float,            # 实际使用的权重
            'arima_forecast': np.array,
            'lstm_forecast': np.array,
            'arima_order': tuple,
            'arima_aic': float,
            'model_fitted': bool,
            'lstm_train_loss': float,
            'lstm_val_loss': float,
            'lstm_epochs_used': int,
            'trend_divergence': bool,         # 预测方向与近120日趋势是否矛盾
            'recent_drift_120d': float,
            'error': str or None,
        }
    """
    # --- 阶段0: ARIMA拟合 ---
    if not STATSMODELS_AVAILABLE:
        return _fallback(log_returns, horizon, "statsmodels not installed")

    try:
        lr = np.array(log_returns, dtype=np.float64)

        # 近期趋势
        recent_120d = lr[-120:] if len(lr) >= 120 else lr
        recent_drift = recent_120d.mean() * 252

        if verbose:
            print("  [1/3] ARIMA拟合...")
        arima_result = fit_arima(lr, order=arima_order, auto_find=auto_find_arima)
        residuals = arima_result['residuals']

        # ARIMA预测
        arima_fc = forecast_arima(arima_result['fitted'], horizon=horizon)
        arima_forecast = arima_fc['forecast']
        arima_drift = arima_fc['annualized_drift']

        if verbose:
            print(f"    ARIMA {arima_result['order']} drift={arima_drift*100:.2f}%")

        # --- 阶段1: LSTM学习残差 ---
        if not TORCH_AVAILABLE:
            if verbose:
                print("    PyTorch不可用，使用纯ARIMA")
            return _arima_only_result(
                arima_forecast, arima_drift, arima_result,
                recent_drift, "PyTorch not installed"
            )

        residuals_float = residuals.astype(np.float32)
        if len(residuals_float) < seq_length * 6:
            if verbose:
                print(f"    残差数据不足({len(residuals_float)})，使用纯ARIMA")
            return _arima_only_result(
                arima_forecast, arima_drift, arima_result,
                recent_drift, f"残差数据不足: {len(residuals_float)}"
            )

        if verbose:
            print("  [2/3] LSTM训练残差修正...")
        lstm_result = _train_lstm_on_residuals(
            residuals_float, seq_length, hidden_dim, num_layers,
            dropout, epochs, verbose
        )

        if not lstm_result['success']:
            return _arima_only_result(
                arima_forecast, arima_drift, arima_result,
                recent_drift, lstm_result.get('error', 'LSTM failed')
            )

        # --- 阶段2: 混合预测 ---
        if verbose:
            print("  [3/3] 生成混合预测...")
        lstm_forecast = forecast_autoregressive(
            lstm_result['model'], lstm_result['last_window'], horizon,
            mean=float(lstm_result['scaler_mean']),
            std=float(lstm_result['scaler_std']),
            clip_std=FORECAST_CLIP_STD
        )

        # 自适应blend_weight
        actual_blend = _compute_blend_weight(
            blend_weight, lstm_result['train_loss'], lstm_result['val_loss']
        )

        combined = arima_forecast + actual_blend * lstm_forecast
        combined_drift = _annualize(combined)
        lstm_correction = combined_drift - arima_drift

        # 趋势方向校验
        trend_divergence = (
            (combined_drift > 0 and recent_drift < 0) or
            (combined_drift < 0 and recent_drift > 0)
        )

        if verbose:
            print(f"    ARIMA: {arima_drift*100:+.2f}%")
            print(f"    LSTM修正: {lstm_correction*100:+.2f}% (weight={actual_blend:.2f})")
            print(f"    混合: {combined_drift*100:+.2f}%")
            if trend_divergence:
                print(f"    ⚠️ 预测方向与近120日趋势({recent_drift*100:+.2f}%)矛盾")

        return {
            'forecast': combined,
            'annualized_drift': combined_drift,
            'arima_drift': arima_drift,
            'lstm_correction': lstm_correction,
            'blend_weight': actual_blend,
            'arima_forecast': arima_forecast,
            'lstm_forecast': lstm_forecast,
            'arima_order': arima_result['order'],
            'arima_aic': arima_result['aic'],
            'model_fitted': True,
            'lstm_train_loss': lstm_result['train_loss'],
            'lstm_val_loss': lstm_result['val_loss'],
            'lstm_epochs_used': lstm_result['epochs_used'],
            'trend_divergence': trend_divergence,
            'recent_drift_120d': recent_drift,
            'error': None,
        }

    except Exception as e:
        return _fallback(log_returns, horizon, str(e))


def _train_lstm_on_residuals(residuals, seq_length, hidden_dim, num_layers,
                              dropout, epochs, verbose):
    """在ARIMA残差上训练LSTM"""
    import torch

    try:
        # 分割
        split_idx = int(len(residuals) * TRAIN_RATIO)
        train_res = residuals[:split_idx]
        val_res = residuals[split_idx:]

        # 标准化（训练集统计量）
        train_norm, val_norm, mean, std = standardize(train_res, val_res)

        # 合并为完整标准化序列
        full_norm = np.concatenate([train_norm, val_norm])

        X, y = prepare_sequences(full_norm, seq_length)
        train_split = split_idx - seq_length
        if train_split < 10:
            return {'success': False, 'error': '训练窗口不足'}

        X_train, y_train = X[:train_split], y[:train_split]
        X_val, y_val = X[train_split:], y[train_split:]

        if len(X_val) < 3:
            return {'success': False, 'error': '验证窗口不足'}

        train_loader = torch.utils.data.DataLoader(
            torch.utils.data.TensorDataset(X_train, y_train),
            batch_size=BATCH_SIZE, shuffle=True
        )
        val_loader = torch.utils.data.DataLoader(
            torch.utils.data.TensorDataset(X_val, y_val),
            batch_size=BATCH_SIZE, shuffle=False
        )

        model = LSTMResidualNet(
            input_dim=1, hidden_dim=hidden_dim, num_layers=num_layers,
            output_dim=1, dropout=dropout
        )

        result = train_lstm(
            model, train_loader, val_loader,
            epochs=epochs, learning_rate=LEARNING_RATE,
            weight_decay=WEIGHT_DECAY, gradient_clip=GRADIENT_CLIP,
            early_stop_patience=EARLY_STOP_PATIENCE,
            lr_patience=LR_PATIENCE, lr_factor=LR_FACTOR,
            min_lr=MIN_LR, verbose=verbose
        )

        return {
            'success': True,
            'model': result['model'],
            'last_window': full_norm[-seq_length:],
            'scaler_mean': mean,
            'scaler_std': std,
            'train_loss': result['train_losses'][-1] if result['train_losses'] else None,
            'val_loss': result['best_val_loss'],
            'epochs_used': result['epochs_used'],
        }

    except Exception as e:
        return {'success': False, 'error': str(e)}


def _compute_blend_weight(user_weight, train_loss, val_loss):
    """计算自适应blend_weight"""
    if user_weight is not None:
        return user_weight

    if train_loss is None or val_loss is None or train_loss <= 0:
        return 1.0

    ratio = val_loss / train_loss
    if ratio > OVERFIT_THRESHOLD_HIGH:
        return min(1.0 / ratio, 0.3)
    elif ratio > OVERFIT_THRESHOLD_LOW:
        return 0.5
    else:
        return 1.0


def _annualize(forecast):
    total_log = forecast.sum()
    total_simple = np.exp(total_log) - 1
    annualized_simple = (1 + total_simple) ** (252 / len(forecast)) - 1
    return np.log(1 + annualized_simple)


def _arima_only_result(arima_forecast, arima_drift, arima_result,
                        recent_drift, error_msg):
    trend_div = (
        (arima_drift > 0 and recent_drift < 0) or
        (arima_drift < 0 and recent_drift > 0)
    )
    return {
        'forecast': arima_forecast,
        'annualized_drift': arima_drift,
        'arima_drift': arima_drift,
        'lstm_correction': 0.0,
        'blend_weight': 0.0,
        'arima_forecast': arima_forecast,
        'lstm_forecast': np.zeros(len(arima_forecast)),
        'arima_order': arima_result['order'],
        'arima_aic': arima_result['aic'],
        'model_fitted': True,
        'lstm_train_loss': None,
        'lstm_val_loss': None,
        'lstm_epochs_used': 0,
        'trend_divergence': trend_div,
        'recent_drift_120d': recent_drift,
        'error': error_msg,
    }


def _fallback(log_returns, horizon, error_msg):
    lr = np.array(log_returns)
    drift = lr.mean() * 252
    recent_drift = lr[-120:].mean() * 252 if len(lr) >= 120 else drift
    return {
        'forecast': np.full(horizon, lr.mean()),
        'annualized_drift': drift,
        'arima_drift': drift,
        'lstm_correction': 0.0,
        'blend_weight': 0.0,
        'arima_forecast': np.full(horizon, lr.mean()),
        'lstm_forecast': np.zeros(horizon),
        'arima_order': (0, 0, 0),
        'arima_aic': float('inf'),
        'model_fitted': False,
        'lstm_train_loss': None,
        'lstm_val_loss': None,
        'lstm_epochs_used': 0,
        'trend_divergence': False,
        'recent_drift_120d': recent_drift,
        'error': error_msg,
    }
