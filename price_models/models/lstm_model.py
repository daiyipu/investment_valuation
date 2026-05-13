"""纯LSTM模型 — 直接对log returns训练LSTM"""

import numpy as np
import pandas as pd
from models.lstm_net import (TORCH_AVAILABLE, LSTMResidualNet, prepare_sequences,
                              create_dataloaders, standardize, train_lstm,
                              forecast_autoregressive)
from config import (SEQ_LENGTH, HORIZON, LSTM_HIDDEN, LSTM_LAYERS, LSTM_DROPOUT,
                    EPOCHS, BATCH_SIZE, LEARNING_RATE, WEIGHT_DECAY, GRADIENT_CLIP,
                    EARLY_STOP_PATIENCE, LR_PATIENCE, LR_FACTOR, MIN_LR,
                    TRAIN_RATIO, FORECAST_CLIP_STD)


def forecast_pure_lstm(log_returns, horizon=HORIZON, seq_length=SEQ_LENGTH,
                       hidden_dim=LSTM_HIDDEN, num_layers=LSTM_LAYERS,
                       dropout=LSTM_DROPOUT, epochs=EPOCHS, verbose=True):
    """纯LSTM模型预测漂移率

    直接对log returns训练LSTM，不经过ARIMA。

    Args:
        log_returns: 对数收益率序列（pd.Series 或 np.array）
        horizon: 预测步数
        seq_length: 输入序列长度
        verbose: 是否打印训练过程

    Returns:
        dict: {
            'forecast': np.array (horizon,),
            'annualized_drift': float,
            'model_fitted': bool,
            'train_loss': float,
            'val_loss': float,
            'epochs_used': int,
            'error': str or None,
        }
    """
    if not TORCH_AVAILABLE:
        return _fallback(log_returns, horizon, "PyTorch not installed")

    try:
        data = np.array(log_returns, dtype=np.float32)
        if len(data) < seq_length * 6:
            return _fallback(log_returns, horizon,
                             f"数据不足: {len(data)} < {seq_length * 6}")

        # 分割 + 标准化
        split_idx = int(len(data) * TRAIN_RATIO)
        train_data = data[:split_idx]
        val_data = data[split_idx:]

        train_norm, val_norm, mean, std = standardize(train_data, val_data)

        # 合并为完整标准化序列用于构建序列
        full_norm = np.concatenate([train_norm, val_norm])

        X, y = prepare_sequences(full_norm, seq_length)
        train_split = split_idx - seq_length
        X_train, y_train = X[:train_split], y[:train_split]
        X_val, y_val = X[train_split:], y[train_split:]

        if len(X_train) < 10 or len(X_val) < 5:
            return _fallback(log_returns, horizon, "训练/验证窗口不足")

        train_loader = __make_loader(X_train, y_train)
        val_loader = __make_loader(X_val, y_val)

        # 创建并训练模型
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

        # 自回归预测
        last_window = full_norm[-seq_length:]
        predictions = forecast_autoregressive(
            result['model'], last_window, horizon,
            mean=float(mean), std=float(std), clip_std=FORECAST_CLIP_STD
        )

        # 年化
        annualized_drift = _annualize(predictions)

        return {
            'forecast': predictions,
            'annualized_drift': annualized_drift,
            'model_fitted': True,
            'train_loss': result['train_losses'][-1] if result['train_losses'] else None,
            'val_loss': result['best_val_loss'],
            'epochs_used': result['epochs_used'],
            'best_epoch': result['best_epoch'],
            'error': None,
        }

    except Exception as e:
        return _fallback(log_returns, horizon, str(e))


def _make_loader(X, y):
    import torch
    dataset = torch.utils.data.TensorDataset(X, y)
    return torch.utils.data.DataLoader(dataset, batch_size=BATCH_SIZE, shuffle=True)


def _fallback(log_returns, horizon, error_msg):
    lr = np.array(log_returns)
    drift = lr.mean() * 252
    return {
        'forecast': np.full(horizon, lr.mean()),
        'annualized_drift': drift,
        'model_fitted': False,
        'train_loss': None,
        'val_loss': None,
        'epochs_used': 0,
        'error': error_msg,
    }


def _annualize(forecast):
    total_log = forecast.sum()
    total_simple = np.exp(total_log) - 1
    annualized_simple = (1 + total_simple) ** (252 / len(forecast)) - 1
    return np.log(1 + annualized_simple)
