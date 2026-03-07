# 市场数据分析模块使用指南

## 概述

本模块用于从Tushare获取上市公司历史交易数据，计算真实的统计指标（波动率、收益率、平均股价等），替代原有的假设参数。

## 模块文件

### 1. 07_market_data_analysis.ipynb
**功能**: 获取历史数据并计算真实指标

**使用步骤**:
1. 设置Tushare Token环境变量
   ```bash
   export TUSHARE_TOKEN='your_token_here'
   ```

2. 配置股票参数
   ```python
   STOCK_CONFIG = {
       'code': '300735.SZ',     # 股票代码
       'name': '光弘科技',       # 股票名称
   }
   ```

3. 运行notebook，自动生成数据文件:
   - `300735_SZ_market_data.json`

**输出指标**:
- **波动率**: 30日/60日/120日/180日滚动波动率
- **收益率**: 不同窗口的年化收益率（漂移率）
- **平均股价**: 不同窗口的移动平均线
- **胜率**: 上涨天数占比
- **价格统计**: 最高价、最低价、标准差等

### 2. utils/market_data_loader.py
**功能**: 加载和使用市场数据

**使用示例**:
```python
from utils.market_data_loader import load_market_data, get_risk_params

# 加载市场数据
market_data = load_market_data('300735.SZ')

# 获取风险参数（推荐使用60日窗口）
risk_params = get_risk_params(market_data, volatility_window='60d')

volatility = risk_params['volatility']  # 真实波动率
drift = risk_params['drift']            # 真实收益率
```

## 数据文件格式

生成的JSON文件包含以下字段：

```json
{
  "stock_code": "300735.SZ",
  "stock_name": "光弘科技",
  "analysis_date": "2024-03-01",

  "current_price": 25.95,
  "avg_price_all": 22.50,

  "volatility_30d": 0.3512,
  "volatility_60d": 0.3117,
  "volatility_120d": 0.2856,
  "volatility_180d": 0.2734,
  "volatility": 0.3117,

  "annual_return_30d": 0.1523,
  "annual_return_60d": 0.0845,
  "annual_return_120d": 0.0567,
  "annual_return_180d": 0.0423,
  "drift": 0.0845,

  "ma_30": 24.56,
  "ma_60": 23.78,
  "ma_120": 22.34,
  "ma_180": 21.89,

  "win_rate_60d": 0.5233,
  "total_days": 450
}
```

## 在其他Notebook中使用

### 方式1: 自动加载（推荐）

更新后的notebook会自动尝试加载市场数据：

```python
try:
    from utils.market_data_loader import load_market_data

    market_data = load_market_data('300735.SZ')

    if market_data:
        volatility = market_data['volatility']
        drift = market_data['drift']
        print(f"使用真实数据: volatility={volatility*100:.2f}%, drift={drift*100:.2f}%")
    else:
        # 回退到默认值
        volatility = 0.35
        drift = 0.08
except:
    # 使用参数文件
    pass
```

### 方式2: 手动加载

```python
from utils.market_data_loader import load_market_data, print_market_data_summary

# 加载数据
market_data = load_market_data('300735.SZ')

# 打印摘要
print_market_data_summary(market_data)

# 使用参数
RISK_PARAMS = {
    'volatility': market_data['volatility'],
    'drift': market_data['drift'],
}
```

## 已更新的Notebook

以下notebook已更新支持真实市场数据：

- ✅ **05_var_calculation.ipynb** - VaR风险测算
  - 自动加载市场数据
  - 使用真实波动率和收益率
  - 如果数据不存在，回退到参数文件

## 推荐工作流程

1. **首次分析新股票**:
   ```bash
   # 1. 运行数据获取脚本
   cd price_maintenance_risk_analysis
   python fetch_gh_data.py

   # 2. 运行市场数据分析
   jupyter notebook notebooks/07_market_data_analysis.ipynb

   # 3. 运行其他风险分析notebook
   jupyter notebook notebooks/05_var_calculation.ipynb
   ```

2. **更新已有股票的数据**:
   - 重新运行 `07_market_data_analysis.ipynb`
   - 其他notebook会自动使用最新数据

3. **切换分析标的**:
   - 修改 `STOCK_CONFIG` 中的股票代码
   - 重新运行相关notebook

## 数据窗口选择建议

| 窗口 | 适用场景 | 说明 |
|------|---------|------|
| 30日 | 短期交易 | 波动较大，反映近期市场情绪 |
| **60日** | **推荐默认** | **平衡短期波动和长期趋势** |
| 120日 | 中期分析 | 更平滑，减少噪音 |
| 180日 | 长期分析 | 最稳定，适合长期投资 |

## 常见问题

### Q: 市场数据文件不存在怎么办？
A: Notebook会自动回退到参数文件中的数据。建议运行 `07_market_data_analysis.ipynb` 生成市场数据。

### Q: 如何更新市场数据？
A: 重新运行 `07_market_data_analysis.ipynb`，会覆盖旧的JSON文件。

### Q: 不同窗口的波动率如何选择？
A:
- 定增锁定期较短（6个月）：推荐60日或90日
- 定增锁定期较长（12+个月）：推荐120日或180日

### Q: 漂移率（drift）是什么？
A: 漂移率是资产的预期年化收益率，基于历史数据计算。详见 `07_market_data_analysis.ipynb` 第6节。

## 技术细节

### 波动率计算
```python
# 年化波动率 = 日收益率标准差 × √252
rolling_std = df['pct_chg'].rolling(window=window).std()
volatility = rolling_std * np.sqrt(252)
```

### 收益率计算
```python
# 年化收益率 = (1 + 期间收益率)^(1/年数) - 1
annual_return = (1 + total_return) ** (1 / years) - 1
```

### 移动平均线
```python
# 简单移动平均
ma = df['close'].rolling(window=window).mean()
```
