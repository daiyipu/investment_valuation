"""LSTM网络定义 — 轻量级残差预测网络"""

import numpy as np

TORCH_AVAILABLE = False
try:
    import torch
    import torch.nn as nn
    TORCH_AVAILABLE = True
except ImportError:
    pass


if TORCH_AVAILABLE:

    class LSTMResidualNet(nn.Module):
        """轻量LSTM网络，用于预测ARIMA残差或直接预测收益率

        约25K参数，CPU训练1-3秒

        Args:
            input_dim: 输入特征维度（默认1）
            hidden_dim: LSTM隐藏层维度（默认32）
            num_layers: LSTM层数（默认2）
            output_dim: 输出维度（默认1）
            dropout: Dropout率（默认0.1）
        """

        def __init__(self, input_dim=1, hidden_dim=32, num_layers=2,
                     output_dim=1, dropout=0.1):
            super().__init__()
            self.hidden_dim = hidden_dim
            self.num_layers = num_layers

            self.lstm = nn.LSTM(
                input_size=input_dim,
                hidden_size=hidden_dim,
                num_layers=num_layers,
                batch_first=True,
                dropout=dropout if num_layers > 1 else 0,
            )
            self.dropout = nn.Dropout(dropout)
            self.fc1 = nn.Linear(hidden_dim, hidden_dim // 2)
            self.relu = nn.ReLU()
            self.fc2 = nn.Linear(hidden_dim // 2, output_dim)

        def forward(self, x):
            # x: (batch, seq_len, input_dim)
            lstm_out, _ = self.lstm(x)
            # 取最后一个时间步的输出
            last_out = lstm_out[:, -1, :]
            out = self.dropout(last_out)
            out = self.relu(self.fc1(out))
            out = self.fc2(out)
            return out

    def prepare_sequences(data, seq_length):
        """将时间序列转为滑动窗口序列

        Args:
            data: numpy array 或 pd.Series
            seq_length: 输入序列长度

        Returns:
            X: (N, seq_length, 1) tensor
            y: (N, 1) tensor
        """
        if isinstance(data, (pd.Series, pd.DataFrame)):
            data = data.values
        data = np.array(data, dtype=np.float32)

        X, y = [], []
        for i in range(len(data) - seq_length):
            X.append(data[i:i + seq_length])
            y.append(data[i + seq_length])

        X = torch.FloatTensor(np.array(X)).unsqueeze(-1)  # (N, seq_len, 1)
        y = torch.FloatTensor(np.array(y)).unsqueeze(-1)   # (N, 1)
        return X, y

    def create_dataloaders(X, y, train_ratio=0.8, batch_size=16):
        """创建训练和验证DataLoader（时间序列顺序分割）

        Returns:
            train_loader, val_loader, split_idx
        """
        split_idx = int(len(X) * train_ratio)

        X_train, y_train = X[:split_idx], y[:split_idx]
        X_val, y_val = X[split_idx:], y[split_idx:]

        train_dataset = torch.utils.data.TensorDataset(X_train, y_train)
        val_dataset = torch.utils.data.TensorDataset(X_val, y_val)

        train_loader = torch.utils.data.DataLoader(
            train_dataset, batch_size=batch_size, shuffle=True
        )
        val_loader = torch.utils.data.DataLoader(
            val_dataset, batch_size=batch_size, shuffle=False
        )

        return train_loader, val_loader, split_idx

    def standardize(train_data, val_data=None):
        """使用训练集统计量标准化

        Returns:
            train_norm, val_norm, mean, std
        """
        mean = train_data.mean()
        std = train_data.std()
        if std < 1e-8:
            std = 1.0

        train_norm = (train_data - mean) / std
        val_norm = (val_data - mean) / std if val_data is not None else None

        return train_norm, val_norm, mean, std

    def forecast_autoregressive(model, last_window, horizon, mean=0, std=1,
                                 clip_std=2.0):
        """自回归预测

        Args:
            model: 训练好的LSTM模型
            last_window: 最近seq_length个值（已标准化）
            horizon: 预测步数
            mean, std: 反标准化参数
            clip_std: 预测值裁剪阈值（±N*std）

        Returns:
            numpy array of denormalized predictions
        """
        model.eval()
        current = np.array(last_window, dtype=np.float32)
        predictions = []

        with torch.no_grad():
            for _ in range(horizon):
                x = torch.FloatTensor(current).view(1, -1, 1)
                pred = model(x).item()
                pred = np.clip(pred, -clip_std, clip_std)
                predictions.append(pred)
                current = np.append(current[1:], pred)

        predictions = np.array(predictions)
        # 反标准化
        predictions = predictions * std + mean
        return predictions

    def train_lstm(model, train_loader, val_loader, epochs=50,
                   learning_rate=0.005, weight_decay=1e-5,
                   gradient_clip=1.0, early_stop_patience=15,
                   lr_patience=10, lr_factor=0.5, min_lr=1e-6,
                   verbose=True):
        """训练LSTM模型

        Returns:
            dict: {
                'model': trained model,
                'train_losses': list,
                'val_losses': list,
                'best_val_loss': float,
                'epochs_used': int,
                'best_epoch': int,
            }
        """
        criterion = nn.MSELoss()
        optimizer = torch.optim.Adam(
            model.parameters(), lr=learning_rate, weight_decay=weight_decay
        )
        scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(
            optimizer, mode='min', factor=lr_factor, patience=lr_patience,
            min_lr=min_lr
        )

        best_val_loss = float('inf')
        best_model_state = None
        patience_counter = 0
        best_epoch = 0
        train_losses = []
        val_losses = []

        for epoch in range(epochs):
            # 训练
            model.train()
            epoch_train_loss = 0
            for X_batch, y_batch in train_loader:
                optimizer.zero_grad()
                output = model(X_batch)
                loss = criterion(output, y_batch)
                loss.backward()
                nn.utils.clip_grad_norm_(model.parameters(), gradient_clip)
                optimizer.step()
                epoch_train_loss += loss.item()

            avg_train = epoch_train_loss / len(train_loader)
            train_losses.append(avg_train)

            # 验证
            model.eval()
            epoch_val_loss = 0
            with torch.no_grad():
                for X_batch, y_batch in val_loader:
                    output = model(X_batch)
                    loss = criterion(output, y_batch)
                    epoch_val_loss += loss.item()

            avg_val = epoch_val_loss / len(val_loader) if len(val_loader) > 0 else float('inf')
            val_losses.append(avg_val)

            scheduler.step(avg_val)

            # 早停
            if avg_val < best_val_loss:
                best_val_loss = avg_val
                best_model_state = {k: v.clone() for k, v in model.state_dict().items()}
                patience_counter = 0
                best_epoch = epoch + 1
            else:
                patience_counter += 1
                if patience_counter >= early_stop_patience:
                    if verbose:
                        print(f"    早停: epoch {epoch+1}, best_val={best_val_loss:.6f}")
                    break

            if verbose and (epoch + 1) % 10 == 0:
                print(f"    epoch {epoch+1}: train={avg_train:.6f}, val={avg_val:.6f}")

        # 恢复最优模型
        if best_model_state:
            model.load_state_dict(best_model_state)

        return {
            'model': model,
            'train_losses': train_losses,
            'val_losses': val_losses,
            'best_val_loss': best_val_loss,
            'epochs_used': len(train_losses),
            'best_epoch': best_epoch,
        }


# 使 pd 在模块级别可用
import pandas as pd
