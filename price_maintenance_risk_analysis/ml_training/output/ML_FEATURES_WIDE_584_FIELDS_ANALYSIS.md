# ml_features_wide表584个字段来源完整分析

## 🎯 直接回答用户的核心问题

### 问题：export_features.py和derive_features.py能否填完584个字段？

**答案：不能完全填满！这两条脚本只能填充部分字段，584个字段来自多个数据源和计算流程。**

## 📊 584个字段的详细来源分析

### 字段分类统计
```
基础字段: 12个
定增相关: 11个  
评分相关: 35个
因子相关: 94个
标签相关: 43个
chip相关: 5个
SUE相关: 8个
其他字段: 376个
总计: 584个字段
```

### 字段来源分解

#### 1. export_features.py 填充的字段（约50-80个）
**数据源**: MySQL数据库直接读取
- **基础信息**: 股票代码、股票简称、报价日、报价日价格、行业分类等
- **财务评分**: 总分_斜率、总分_趋势、盈利能力_斜率、盈利能力_趋势等
- **定增参数**: 定增_发行价、定增_增发数量、定增_募资总额、定增决策等
- **PB/PE/PS**: 个股PB、个股PE、个股PS、个股市值
- **chip指标**: chip_winner_rate、chip_avg_cost_dev等
- **FCF数据**: FCF_T、FCF_T1、FCF年份_T等（约16个字段）
- **SUE数据**: sue_yoy、sue_zscore、sue_beat等（约8个字段）

#### 2. derive_features.py 填充的字段（约100-150个）
**计算方式**: 从基础特征派生
- **FCF增长率**: FCF增长率相关字段（约18个）
- **动量波动率**: 波动率_20d、年化收益_60d、区间收益_120d等（约12个）
- **行业相对估值**: 相对行业估值的指标（约8个）
- **市场相对表现**: 相对市场的指标（约6个）
- **时间衰减**: 时间衰减相关指标（约4个）
- **行业市值占比**: 行业市值占比相关（约3个）
- **市场指数相对**: 市场指数相对指标（约2个）
- **Beta因子**: beta_mkt_60、beta_mkt_120、beta_ind_120等（约10个）
- **技术指标**: RSI、KDJ、MACD、布林带等（约20个）
- **量价指标**: vol_ratio_20、amount_ratio_20、corr_close_vol_20等（约15个）
- **滚动矩**: skew_20、kurt_60、roc_5等（约12个）

#### 3. factor_engine.py 计算的字段（约40-50个）
**计算方式**: 通过factor_engine.py算子库计算
- **Beta族**: beta_mkt_{60,120,250}、beta_ind_120、idiovol_120
- **Alpha158族**: K线形态、技术指标、量价等因子
- **特殊因子**: 通过tushare API计算的因子

#### 4. 其他脚本填充的字段（约200-300个）
**来源**: 以下脚本和流程
- **build_backtest_panel.py**: 回测面板构建，填充大量行情和市场因子
- **fetch_factors.py**: 因子数据获取
- **eval_factors.py**: 因子评估时计算的指标
- **数据回填脚本**: 各种特征回填脚本
- **历史数据导入**: 从历史数据导入的字段

## 🔍 发现的双轨数据架构

### 轨道1: 结构化存储（ml_features_wide表）
```
ml_features_wide表 (584个字段，568,624个样本)
    - 用于ML训练
    - 结构化存储
    - 支持SQL查询
```

### 轨道2: 面板文件（parquet文件）
```
backtest_panel.parquet (回测面板)
features.parquet (训练特征)
features_derived.parquet (衍生特征)
    - 用于回测分析
    - 文件存储
    - 便于传输和备份
```

## 💡 关键发现

### 1. 并非单一脚本填充
**584个字段不是由一个脚本填充的**，而是由多个脚本和流程共同构建：

- **export_features.py**: 填充约50-80个基础字段
- **derive_features.py**: 填充约100-150个衍生字段
- **factor_engine.py**: 计算约40-50个因子字段
- **其他流程**: 填充约200-300个其他字段

### 2. 两条并行路径
**路径A: ML训练路径**
```
MySQL数据源 → export_features.py → ml_features_wide表 → ML训练
```

**路径B: 回测分析路径**
```
MySQL数据源 → build_backtest_panel.py → backtest_panel.parquet → 回测分析
```

### 3. ml_features_wide表的复杂构建过程
ml_features_wide表的584个字段是**逐步构建**的，不是一次性完成的：

1. **初始创建**: 基础字段（股票代码、报价日等）
2. **数据导入**: 从MySQL导入的基础数据
3. **特征计算**: export_features.py填充财务等特征
4. **因子计算**: derive_features.py + factor_engine.py计算因子
5. **历史回填**: 从历史数据补充字段
6. **特殊处理**: 各种特殊脚本处理的字段

## 🎯 结论

### 回答用户的核心问题

**Q: export_features.py和derive_features.py能否填完584个字段？**

**A: 不能！这两个脚本只能填充约150-230个字段（约25-40%），剩余的字段来自：**
- factor_engine.py计算的因子
- build_backtest_panel.py构建的回测特征
- 其他各种特征获取和计算脚本

### 完整的字段填充路径

```
ml_features_wide表584个字段来源：

1. export_features.py: ~15% (50-80字段)
   └─ 从MySQL读取基础数据

2. derive_features.py: ~25% (100-150字段)
   └─ 从基础特征派生

3. factor_engine.py: ~8% (40-50字段)
   └─ 因子算子库计算

4. 其他脚本和流程: ~52% (200-300字段)
   └─ 各种历史数据、回测特征、特殊计算等
```

### 建议

**如果要完整使用ml_features_wide表**：
1. 先运行export_features.py填充基础字段
2. 再运行derive_features.py填充衍生字段  
3. 通过其他脚本补充剩余字段
4. 或者使用build_backtest_panel.py一次性构建完整面板

**factor_engine.py的作用**：
- 提供标准化的因子计算算子
- 被derive_features.py调用
- 计算结果通过derive_features.py的输出返回
- 本身不直接写入数据库，而是作为计算函数库

---

**总结**: 584个字段是由多个脚本和流程共同构建的，export_features.py和derive_features.py只能填充约40%的字段，剩余字段来自其他数据源和计算流程。