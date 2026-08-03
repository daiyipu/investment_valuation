# 新工程失败的根本原因分析（纠正版）

## ❌ **我之前的分析是错误的！**

### 错误观点
之前我说"新工程缺乏独立数据基础" - **这是错误的**！

### 正确事实
新工程**确实也用了 ml_features_wide 表**，数据来源和老项目一样。

---

## 🔍 **新工程失败的真正原因**

### 1. **❌ 特征工程简化过度**

#### 老项目的专业特征工程
```python
# 老项目 train_scorecard_model.py
def train_scorecard_model():
    # 1. 五步特征选择漏斗
    selected = select_features(X, y, X_test)  # IV→PSI→相关→VIF→LGB
    # 2. WOE分箱
    bins = fit_woe(X_train, y_train, selected_features)
    X_train_woe = apply_woe(X_train, bins)
    # 3. 逻辑回归
    model = LogisticRegression().fit(X_train_woe, y_train)
```

#### 新工程的简化特征工程
```python
# 新工程 placement_only_pipeline.py
def prepare_ml_data(df):
    # 简单排除模式
    exclude_patterns = ['报价日', '股票代码', ...]
    feature_cols = [c for c in df.columns if not any(p in c for p in exclude_patterns)]
    X = df[feature_cols].fillna(0)  # 简单填充

def train_and_validate():
    # 简单相关性特征选择
    correlations = X.corrwith(y).abs().sort_values(ascending=False)
    top_features = correlations.head(20).index.tolist()  # 只选Top 20
    # 直接逻辑回归，没有WOE分箱
    model = LogisticRegression().fit(X_train[top_features], y_train)
```

**差异**：
- ❌ 新工程缺少五步特征选择漏斗
- ❌ 新工程没有WOE分箱处理
- ❌ 新工程特征数量固定20个，不是精选的18个

### 2. **❌ 验证流程错误**

#### 老项目的正确验证
```python
# 老项目 backtest_long_short.py
def run(model_ver, horizon, sample_type='fake_quote'):
    # 明确使用全A数据，排除定增样本
    sql = """
    SELECT * FROM ml_features_wide
    WHERE sample_type='fake_quote'  # 明确排除placement
    AND `7个月涨跌幅` IS NOT NULL
    """
    panel = pd.read_sql(sql, conn)

    # 逐截面打分 + 多空回测
    for date, group in panel.groupby('报价日'):
        scores = score_sc(model, group[model_features])
        # Top/Bottom 10%等权
        long_ret = group[scores > percentile_90].mean()
        short_ret = group[scores < percentile_10].mean()
```

#### 新工程的错误验证
```python
# 新工程 placement_only_pipeline.py
def train_and_validate():
    # ❌ placement数据内部随机分割
    X_train, X_val, y_train, y_val = train_test_split(
        X_selected, y, test_size=0.2, random_state=42, stratify=y
    )
    # 这不是真正的全A验证！
```

**差异**：
- ❌ 新工程用placement数据内部分割，不是全A验证
- ❌ 新工程没有Long/Short回测
- ❌ 新工程没有计算IC/ICIR

### 3. **❌ 模型训练简化**

#### 老项目的专业训练
```python
# 老项目 train_scorecard_model.py
# 1. scorecardpy分箱
import scorecardpy as sc
bins = sc.woebin(df, y='y', method='chimerge', max_bin=5)

# 2. WOE变换
for feat in features:
    X_train[f'{feat}_woe'] = apply_woe(X_train[feat], bins[feat])

# 3. 逻辑回归 on WOE
model = LogisticRegression().fit(X_train_woe, y_train)

# 4. 模型入库 + 注册为生产模型
save_model_meta(model, metadata)
register_version(model_id, 'current.full')
```

#### 新工程的简化训练
```python
# 新工程 placement_only_pipeline.py
# 1. 简单特征选择
top_features = correlations.head(20).index.tolist()

# 2. 直接逻辑回归（无WOE）
model = LogisticRegression(C=1.0, penalty='l2', max_iter=1000)
model.fit(X_train[top_features], y_train)

# 3. 简单AUC评估
val_pred = model.predict_proba(X_val[top_features])[:, 1]
val_auc = roc_auc_score(y_val, val_pred)
```

**差异**：
- ❌ 新工程没有WOE分箱变换
- ❌ 新工程没有scorecardpy专业分箱
- ❌ 新工程没有模型入库和版本管理

### 4. **❌ 标签处理不当**

#### 老项目的正确标签
```python
# 老项目中，标签已经正确存储在ml_features_wide表中
SELECT `标签_盈利_-10_7m` FROM ml_features_wide
# placement: 正样本率 57.8%
# fake_quote: 正样本率 44.3%
```

#### 新工程的重新计算标签
```python
# 新工程 full_pipeline_with_labels.py
# 重新计算标签（可能导致不一致）
batch_df['标签_盈利_-10_7m'] = (
    pd.to_numeric(batch_df['7个月涨跌幅'], errors='coerce') > -10
).astype(int)
```

**差异**：
- ⚠️ 新工程重新计算标签，可能与老项目不一致
- ⚠️ 新工程缺少标签质量检查

---

## 📊 **新工程 vs 老项目详细对比**

| 维度 | 老项目（成功） | 新工程（失败） | 关键差异 |
|------|----------------|----------------|----------|
| **数据来源** | ml_features_wide | ml_features_wide | ✅ 相同 |
| **特征选择** | 五步漏斗 | 简单相关性Top20 | ❌ **致命差异** |
| **WOE分箱** | scorecardpy专业分箱 | 无WOE分箱 | ❌ **致命差异** |
| **模型训练** | LR on WOE特征 | LR on 原始特征 | ❌ **致命差异** |
| **验证方式** | 全A fake_quote | placement内部分割 | ❌ **致命差异** |
| **回测分析** | Long/Short + IC/ICIR | 简单AUC | ❌ **致命差异** |
| **模型管理** | 入库 + 版本注册 | 无版本管理 | ❌ 管理差异 |

---

## 🎯 **新工程失败的4个致命问题**

### 1. **❌ 特征工程简化过度**
```python
# 老项目: 400+特征 → IV筛选 → PSI筛选 → 相关性筛选 → VIF筛选 → LGB剪枝 → 18特征
# 新工程: 400+特征 → 简单相关性 → Top 20特征
```

### 2. **❌ 缺少WOE分箱处理**
```python
# 老项目: scorecardpy.woebin → WOE变换 → 逻辑回归
# 新工程: 原始特征 → 逻辑回归（无WOE）
```

### 3. **❌ 验证流程错误**
```python
# 老项目: placement训练 → fake_quote验证 → Long/Short回测
# 新工程: placement训练 → placement内部分割验证
```

### 4. **❌ 缺少专业回测分析**
```python
# 老项目: IC/ICIR + 年化收益 + 夏普比率 + 月度/年度分解
# 新工程: 简单AUC + KS
```

---

## 💡 **结论**

### 新工程为什么失败？

**不是因为数据表**（新工程也用了ml_features_wide）
**而是因为**：
1. ❌ 特征工程简化过度（缺少五步漏斗和WOE分箱）
2. ❌ 验证流程错误（placement内部分割，不是真正的全A验证）
3. ❌ 模型训练简化（没有WOE变换，直接用原始特征）
4. ❌ 缺少专业回测（没有Long/Short策略和IC/ICIR分析）

### 新工程如何改进？

**需要复用老项目的成功经验**：
1. ✅ 实现五步特征选择漏斗
2. ✅ 加入scorecardpy WOE分箱
3. ✅ 使用正确的全A验证流程
4. ✅ 加入Long/Short回测和IC/ICIR分析

---

**纠正时间**: 2026-07-03 22:00:00
**纠正原因**: 用户正确指出新工程也用了ml_features_wide表
**新结论**: 新工程失败的原因不是数据表，而是特征工程和验证流程的简化过度
