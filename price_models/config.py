"""模型参数配置"""

# 数据参数
DATA_WINDOW = 250          # 使用最近250个交易日
SEQ_LENGTH = 20            # LSTM输入序列长度（约1个月）
TRAIN_RATIO = 0.8          # 训练集比例（时间序列分割）
HORIZON = 120              # 预测期数（约半年）

# LSTM网络结构
LSTM_HIDDEN = 32           # 隐藏层维度
LSTM_LAYERS = 2            # LSTM层数
LSTM_DROPOUT = 0.1         # Dropout率

# 训练参数
EPOCHS = 50                # 最大训练轮数
BATCH_SIZE = 16            # 批大小
LEARNING_RATE = 0.005      # 初始学习率
WEIGHT_DECAY = 1e-5        # L2正则化
GRADIENT_CLIP = 1.0        # 梯度裁剪
EARLY_STOP_PATIENCE = 15   # 早停耐心值
LR_PATIENCE = 10           # 学习率衰减耐心值
LR_FACTOR = 0.5            # 学习率衰减因子
MIN_LR = 1e-6              # 最小学习率

# 混合模型参数
BLEND_WEIGHT = 1.0         # LSTM修正权重（默认全量使用）
OVERFIT_THRESHOLD_HIGH = 2.0   # val_loss/train_loss > 此值，blend=1/ratio
OVERFIT_THRESHOLD_LOW = 1.5    # val_loss/train_loss > 此值，blend=0.5
FORECAST_CLIP_STD = 2.0    # LSTM预测裁剪（±N倍标准差）

# DB路径
DB_PATH = None             # None表示使用默认路径
