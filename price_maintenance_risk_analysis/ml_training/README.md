# ml_training — 定增盈利预测模型(训练/评估/预测)

基于历史定增(2010–2026, ~6400 样本,东方财富源)训练评分卡/LightGBM 模型,预测锁定期后盈利概率。
**真相源**:`placement_evaluation` MySQL 表(标签 + 原始因子);特征由 `export_features` 统一发射。
**当前生产**:7m 灰度评分卡(`current.full`,共识特征,LOYO AUC ~0.62–0.73);附 1m/3m 灰度 SC 辅助列。

> ⚠️ 评估一律用 **LOYO(留一年出)**,不用 shuffle 5折CV(同股近邻日自相关致虚高 +0.1~0.25)。

---

## 三层数据架构

```
L1 取数 → scripts/data_pipeline/        (fetch_factors/compute_labels/...  → placement_evaluation DB)
L2 发射 → ml_training/export_features.py (DB → data/features.parquet, 统一发射全族+标签)
L3 衍生 → ml_training/derive_features.py (features.parquet → features_derived.parquet)
   训练 → train_to_production / train_scorecard_model (→ DB ml_model_meta + model_registry.json)
   预测 → predict_profitability.py (scored Excel → 7m SC + BLUE + 1m/3m 辅助列)
```
一键全链路:`python scripts/data_pipeline/update_all_data.py`(12 步, `--from N` 续跑)。

---

## 目录结构(按职责分组)

### 特征层(L2/L3)
| 文件 | 职责 |
|---|---|
| `export_features.py` | **L2 发射器**:`SELECT * placement_evaluation` → `features.parquet`。统一发射所有族(行情/估值/FCF/定增结构/chip/资金流 mf·北向 nb/SMC smc_*/标签 盈利·极性·超·短)。gray 标签(超额 p75/p25)在此算。 |
| `derive_features.py` | **L3 衍生器**:`features.parquet` → `features_derived.parquet`。FCF增长率/交叉/评分delta/估值相对/行情动量/行业·大盘/策略信号(波二·抵抗)/因子引擎(Beta+Alpha158+SMC,需tushare)。每族 try/except 独立失败不影响其余。 |
| `factor_engine.py` | OHLCV → 因子库:K线/技术/量价/动量/Beta + 多周期(W/M) + **SMC 聪明钱**(`smc_factors` 日线 + `smc_factors_multiperiod` 周月线)。predict 端自动调(`derive_alpha_beta_factors`)。 |
| `feature_exclusions.py` | **剔除清单单一源**:`LEAKAGE_PATTERNS`(报价日/涨跌幅/周涨跌幅/excess_/标签 等)+ `BUSINESS_DROP`(绝对水平/常量)。`get_excluded_columns()` 全项目共用。 |

### 选择 + 评估
| 文件 | 职责 |
|---|---|
| `feature_selection.py` | **canonical 五步固定**:`IV>0.01 → PSI≤0.25 → |r|≤0.7 → VIF<5 → LGBM剪`。阈值常量在顶部(改阈值只改此处)。`select_features()`+`prune_by_lgb_importance()`。 |
| `eval_loyo.py` | **LOYO 验证**(去偏):`fit_woe`/`apply_woe`/`loyo_fixed`(锁定特征跨折 mean±std)。评分卡标准评估。 |
| `eval_factors.py` | **单因子效力**(`python eval_factors.py <因子> [--prefix X_]`):coverage-aware + **粒度自动对齐**(日→短标签1w/2w/4w, 周→1m/3m, 月→3m/7m)。挖新因子先用它评。 |
| `validate_methods.py` | AUC/KS/IV 指标 + `run_part_a` OOT 切分(train≤Y/test>Y)。 |
| `validate_model.py` | 外部测试集验证(旧, 低覆盖参考)。 |

### 训练
| 文件 | 职责 |
|---|---|
| `train_to_production.py` | **生产标准入口**:LOYO 共识特征(跨折频次≥N)→ 锁定特征 LOYO → 部署 SC/LGB(`--set-current`)。7m 生产模型由此产出。 |
| `train_scorecard_model.py` | WOE 评分卡训练 → DB(`run()`, 含 LOYO 去偏自评 + meta 存 `sc_loyo_auc` 真值)。 |
| `train_models.py` | LGB+LR 双模型训练(`--label`/`--split-year`/`--no-set-current`)。 |
| `train_scorecard.py` | WOE/IV 底层(`calc_iv_all_features`/`fit_woe` 等,被各处复用)。 |
| `train_horizon_models.py` | 各期限独立特征集模型(全 set_current=False, 扫描用)。 |
| `compare_selection.py` | 特征选择方法对比。 |

### 模型管理
| 文件 | 职责 |
|---|---|
| `db_model_store.py` | 模型**元信息+权重存 DB**(`ml_model_meta` 表)。`load_predict_bundle`/`save_model_meta`/`list_model_metas`。predict 从 DB 读权重。 |
| `model_registry.py` | **版本指针**(`output/model_registry.json`):`current.full`/`current.scorecard` 指向生产版本。`get_current`/`set_current`。 |
| `manage_models.py` | CLI:`list`/`set <type> <version>`(切生产)/`current`。 |
| `db_dataset_store.py` | 数据集快照(`ml_dataset_snapshot` 表, parquet BLOB, 按 sha256 去重)。防 qfq 漂移致不可复现。 |
| `manage_snapshots.py` | CLI:`list`/`restore <version>`/`show`。 |

### 预测 + 解释 + 诊断
| 文件 | 职责 |
|---|---|
| `predict_profitability.py` | **预测主入口**:`predict(scored_excel)`。输出 7m 评分卡(主)+ BLUE(同期对比)+ 1m/3m 辅助列。predict 端实时算 SMC(derive_alpha_beta_factors→factor_engine)。 |
| `explain_scorecard.py` | 评分卡概率**拆解**(只读):每特征 logit 贡献, 找扣分项。 |
| `diagnose_bin_drift.py` / `diagnose_strategy_scorecard.py` | 诊断(分箱漂移/策略评分卡)。 |

### 数据
| 路径 | 内容 |
|---|---|
| `data/features.parquet` | L2 基础特征(DB 发射, ~290 列)。gitignore, 可重建。 |
| `data/features_derived.parquet` | L3 衍生后(~390 列, 含 SMC/策略/多周期)。训练输入。gitignore。 |
| `data/features_schema.md` | 字段说明(跟踪)。 |
| `output/` | 训练产物(model_registry.json 跟踪; v_*/pkl/lgb 等 gitignore, 存 DB)。 |

---

## 常用操作

### 加一个新因子
1. **计算**:`factor_engine.py` 加函数(若 OHLCV 派生)或 `scripts/data_pipeline/fetch_factors.py` 加子命令(若 tushare 取数)→ 写 `placement_evaluation` DB。
2. **发射**:`export_features.py` 加 `r['新因子']=pe.get(...)`(若 smc_* 走通用前缀循环则免改)。
3. **评估**:`python ml_training/eval_factors.py 新因子`(coverage-aware + 粒度对齐, IV/AUC)。
4. **入选**:若 IV>0.01 且 LOYO 不退化,canonical 选择会自动纳入。重训 `train_to_production.py`。

### 训练生产模型(7m)
```bash
cd price_maintenance_risk_analysis
python scripts/data_pipeline/train_to_production.py ml_training/data/features_derived.parquet \
       --horizon 7 --kind gray --model sc --set-current
```
(共识特征 LOYO + 部署。`--set-current` 切生产; 不加则 set_current=False。)

### 评估因子效力
```bash
python ml_training/eval_factors.py smc_premium_discount chip_concentration   # 指定
python ml_training/eval_factors.py --prefix smc_                              # 全 smc_*
python ml_training/eval_factors.py <因子> --horizons all                      # 不对齐, 全期限
```

### 预测 / 解释
```bash
python scripts/score_one_stock.py 300604.SZ 长川科技 20260618      # 单股终端表格
python scripts/batch_screen_and_score.py --input <名单.xlsx> --sheet 0  # 批量(评分+ML)
python ml_training/explain_scorecard.py 300604.SZ 长川科技 20260618      # 评分卡拆解
```

### 切换 / 查看生产模型
```bash
python ml_training/manage_models.py current
python ml_training/manage_models.py list
python ml_training/manage_models.py set full <version>   # 回滚生产
```

---

## 关键约定(必读)

1. **评估用 LOYO,不用 shuffle CV** — shuffle 把同股近邻日混入 train/test 致虚高。`eval_loyo.loyo_fixed` / `validate_methods.run_part_a` 是去偏标准。
2. **coverage-aware 评估** — 历史有限因子(chip 2018+/北向 2016+/moneyflow 2013+)只在自身非空样本评,不 median 填 pooled(否则假阴性)。`eval_factors.py` 已内置。
3. **粒度对齐** — 因子 bar 粒度匹配 target 期限(日→短标签 1w/2w/4w, 周→1m/3m, 月→3m/7m)。错配=假阴性。
4. **canonical 特征选择唯一** — `feature_selection.py` 五步固定(IV>0.01→PSI→corr→VIF→LGBM),不再各脚本各搞一套。
5. **标签防泄漏** — `feature_exclusions` 拦截 涨跌幅/周涨跌幅/excess_/标签/报价日 等。新增目标变量列务必加进排除清单。
6. **token 不硬编码** — 走 `tushare_token.resolve_tushare_token()`(旧 literal 已泄漏需轮换)。
7. **模型权重在 DB**(`ml_model_meta`),版本指针在 `model_registry.json`(`current.full`)。
8. **set_current=False 扫描** — 多期限/实验模型一律不切生产,避免静默切换 regime;生产提升经 `manage_models.py set` 人工确认。

---

## 关键发现与边界(2026-06 因子挖掘结论)

- **模型本质 = regime/动量(beta),非选股 α**。超额收益标签(剥 regime)实测 AUC 反降 → 现有动量特征赚的就是 regime 的钱。
- **AUC 天花板 ~0.65–0.73**(LOYO)。定增专属特征多已定价;非公开筹码类(chip)有信号但受历史覆盖陷阱。
- **期限可预测性**:月级 0.63–0.67(中),**周级 ~0.55(近噪声)**。短期收益=噪声主导。
- **聪明钱(SMC)**:仅 `premium_discount`(区间位置, 反向均值回归)微弱有效(IV 0.035–0.054);ote/bos/fvg/displacement/sweep ≈0。周/月线 bar 配月标签是关键(日线配月标签错配)。
- **有信号的因子**:`chip_concentration`(2018+ 子集 +2.5pp, 覆盖陷阱)、`smc_premium_discount`、技术/动量族。
- **无效因子**:定增结构条款(公开已定价)、SMC 多数变体、moneyflow(94%覆盖仍~0)、超额收益标签。

---

## 依赖

- Python ≥ 3.10(用 vnpy 环境:`~/anaconda3/envs/vnpy/bin/python`, 非 base py3.7)
- lightgbm, scikit-learn, pandas, numpy, pymysql, openpyxl, tushare
- MySQL(investment_valuation 库:placement_evaluation/market_data/relative_valuation/historical_fcf/industry_daily/industry_data/issue_date_locked/ml_model_meta/ml_dataset_snapshot 等)
