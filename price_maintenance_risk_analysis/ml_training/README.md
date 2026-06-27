# ml_training — 定增盈利预测模型(训练/评估/预测)

基于历史定增(2010–2026, ~6400 样本,东方财富源)训练评分卡/LightGBM 模型,预测锁定期后盈利概率。
**真相源**:`placement_evaluation` MySQL 表(标签 + 原始因子)+ `ml_train_wide`(**模型训练宽表**,训练/验证样本空间唯一来源)。
**当前生产**:7m 灰度评分卡(`current.full`,共识特征,LOYO AUC ~0.62–0.73);附 1m/3m 灰度 SC 辅助列。

> ⚠️ 评估一律用 **LOYO(留一年出)**,不用 shuffle 5折CV(同股近邻日自相关致虚高 +0.1~0.25)。

---

## 三层数据架构(★三层一入口铁律)

```
L0 取数 → ml_training/data/         (ingest_raw/fetch_factors/labels/industry/financial/market → DB 时序表)
L1 装配 → ml_training/features/     (build_features(code,date) = 唯一特征装配入口; 定增/回测/predict 三产线共用)
L2 训练 → ml_training/train/        (train_to_production / train_scorecard_model → DB ml_model_meta)
L3 验证 → ml_training/validate/     (backtest_long_short / eval_loyo → DB ml_validation*)
L4 部署 → ml_training/deploy/       (predict_profitability / score_one_stock / batch_*)
```
**数据源**:训练/验证样本从 `ml_train_wide`(DB 宽表)`load_panel(tag)` 直读 DataFrame(无 parquet 文件)。
一键取数:`python ml_training/data/update_all_data.py`。

---

## ★ ml_train_wide — 模型训练宽表(DB 样本空间唯一来源)

全特征+全标签的宽表,训练/验证的基础。train/val/test 子集都从此切分(内存按年份/随机/分层),不另建零散样本表。新字段/新标签统一维护其上。

| tag | 内容 | 用途 |
|---|---|---|
| `placement_train_20260627` | 定增 5739 样本 × 305 列(含定增标签+最终结论) | **SC 模型训练源** |
| `fullA_5labels_20260627` | 全A 5600股 × 414特征 + 5标签 × 192月截面(563073行) | **回测/验证源** |

**入口**(`ml_training/validate/save_validation_db.py`):
- `load_panel(tag)` / `load_features(src)` — DB BLOB → DataFrame 直读(训练/LOYO 用,无文件)。
- `register_panel(path, tag)` — panel 入库(sha256 去重,入库后删源文件)。
- `restore_panel(tag, out_path)` — DB→parquet 文件(仅外部工具需文件时)。
- `list_panels()` — 列所有样本空间。

前置:MySQL `max_allowed_packet=1G`(`/usr/local/etc/my.cnf` + `brew services restart mysql`)。

---

## 目录结构(按 7 段分,重构后)

### samples/ — 样本脊柱(代码,回溯日期)
| 脚本 | 职责 |
|---|---|
| `fetch_universe.py` | 全A universe 构建(多空回测用),PIT 在市过滤。 |
| `fetch_placements.py` | 抓近10年A股定增名单+日期(东方财富 datacenter)。 |
| `gen_backtest_samples.py` | 生成回测样本清单:(股票,月末报价日)×universe×月,PIT 过滤。 |
| `build_pilot_samples.py` | 建 pilot universe:N股×192月末。 |

### data/ — 取数落库(L0)
| 脚本 | 职责 |
|---|---|
| `ingest_raw.py` | **统一 L0 摄入入口**:逐股 ingest_stock_full(market+FCF+行业)→ DB 时序表。 |
| `update_all_data.py` | **一键全链路编排**(指数→行情→定增→标签→因子→特征→衍生,`--from N` 续跑)。 |
| `labels.py` | 标签计算(合并 compute_labels+recompute_label_qfq+backfill_evaluations);子命令 `labels\|qfq\|evaluations`。 |
| `industry.py` | 行业取数(合并 refresh_industry_daily+mapping+fetch_em_industry_stocks);`refresh\|mapping\|em`。 |
| `financial.py` | 财务取数(合并 batch_financial_score+backfill_financial_indicators+bulk_backfill_early_scores+backfill_ann_date);`score\|indicators\|early\|ann_date`。 |
| `update_market_data.py` | **个股行情/FCF/行业全量摄入**(核心, `ingest_stock_full`)。 |
| `backfill_market_data.py` | 行情表缺口补全(qfq/daily_basic/monthly 到≥99%)。 |
| `update_indices_data.py` | 市场指数数据(沪深300等 + locked_date 指数统计)。 |
| `fetch_factors.py` | 统一因子摄入 → placement_evaluation DB(chip/mf/smc/nb/sue)。 |

> data/ 也放 parquet 数据文件(features_backtest_*/samples_*/sue_timelines 等,gitignored,可重建)。panel 复制品已删,canonical 在 ml_train_wide。

### features/ — 特征工程(★一处维护全特征)
| 脚本 | 职责 |
|---|---|
| `build_backtest_panel.py` | **★ build_features(samples) 唯一装配入口**(三层一入口·组装层):prefetch→基列(FCF/评分/比率)→run_derivation→specials。定增/回测/predict 三产线共用。 |
| `derive_features.py` | **派生核心**(`run_derivation` + 11 个 derive_*):FCF增长/交叉/评分delta/估值相对/行情动量/行业·大盘/策略信号/alpha_beta因子/idx_factor。PIT ≤报价日。 |
| `export_features.py` | 定增特征导出(placement):load_scored + build_features → 训练 panel;PIT loaders(load_fcf_bulk/load_financial_ratios/load_specials/load_pb_vs_industry 等)。 |
| `factor_engine.py` | 因子引擎(声明式算子库):K线/技术/量价/动量/Beta+Alpha158+多周期(W/M)。 |
| `feature_selection.py` | **canonical 五步特征选择唯一入口**:IV→PSI→corr→VIF→LGBM。 |
| `feature_exclusions.py` | 特征剔除清单(LEAKAGE/BUSINESS)+ `assert_panel_superset` 护栏。 |

### train/ — 训练(LOYO)
| 脚本 | 职责 |
|---|---|
| `train_to_production.py` | **生产标准入口**:LOYO共识特征→锁定特征LOYO→部署SC/LGB(`--set-current`)。7m主模型由此产出。 |
| `train_scorecard_model.py` | WOE评分卡训练→DB(`run()`,含LOYO去偏自评+meta存sc_loyo_auc)。 |
| `train_scorecard.py` | WOE/IV底层(`calc_iv_all_features`/`woe_transform`,被各处复用)。 |
| `train_horizon_models.py` | 各期限独立特征集模型(全 set_current=False,扫描用)。 |
| `train_models.py` | LGB+LR 双模型训练(旧版基线)。 |
| `sweep_label.py` | 标签阈值扫描(双样本空间:非灰AUC/KS+全样本IC+灰度%)。 |

### validate/ — 验证(含评估底层引擎)
| 脚本 | 职责 |
|---|---|
| `eval_loyo.py` | **LOYO 标准评估引擎**:`fit_woe`/`apply_woe`(CV 拟合-应用)/`run`(逐年留一)。项目唯一评估口径,被 train/report 共用。 |
| `validate_methods.py` | **验证指标+打分应用**:AUC/KS/IV + `apply_woe_score`(打分应用WOE)+`make_features`。被 train/report/scoring 共用。 |
| `backtest_long_short.py` | **全A多空回测+IC/ICIR**(量化核心验证);`fwd_returns`/`eval_cross_section`。 |
| `save_validation_db.py` | **验证/报告入库唯一入口**:`save_validation_run`→ml_validation*;`save_model_report`→ml_model_report;`load_panel`/`register_panel`/`restore_panel`/`list_panels`(ml_train_wide)。 |
| `validate_5h.py` | 5期限验证(补return_3m/1m/1w/2w标签+IC/ICIR/L-S)。 |
| `eval_factors.py` | 单因子效力(coverage-aware+粒度对齐)。 |
| `validate_model.py` | 外部测试集验证(旧,低覆盖参考)。 |

### report/ — 报告
| 脚本 | 职责 |
|---|---|
| `report_horizon.py` | 任意期限gray SC模型LOYO per-year报告(AUC/KS/IC)。 |
| `audit_scorecard.py` | 评分卡审计(上线前人审:每特征业务解释/IV/系数/分箱/分值)。 |
| `explain_scorecard.py` | 评分卡概率拆解(单股为何得分高/低,只读)。 |
| `compare_lgb_sc.py` | LGB vs SC × 共识/全量字段对比。 |
| `compare_selection.py` | 特征选择方法对比(IV vs Lasso vs 逐步回归)。 |
| `feature_glossary.py` | 特征业务解释词典(审计/核查基础)。 |

### deploy/ — 部署
| 脚本 | 职责 |
|---|---|
| `predict_profitability.py` | **预测主入口**:`score_sc`(模块级生产打分)+`predict`(7m SC+BLUE+1m/3m辅助列)。 |
| `score_one_stock.py` | 单股定增评分(终端表格输出)。 |
| `batch_screen_and_score.py` | 批量定增筛选+财务评分一体化。 |
| `batch_screener.py` | 批量筛选器。 |
| `db_model_store.py` | 模型元信息+权重存DB(`ml_model_meta`):`load_predict_bundle`/`save_model_meta`。 |
| `db_dataset_store.py` | 数据集快照(`ml_dataset_snapshot`,parquet BLOB,sha256去重)。 |
| `model_registry.py` | 版本指针(`model_registry.json`):`current.full`/`current.scorecard`。 |
| `manage_models.py` | CLI:`list`/`current`/`set`(切生产)/`info`。 |
| `manage_snapshots.py` | CLI:`list`/`show`/`restore`。 |

### pipeline/ — 评估底层(LOYO/WOE 引擎)
| 脚本 | 职责 |
|---|---|
| `eval_loyo.py` | **LOYO 标准评估**:`fit_woe`/`apply_woe`(CV 拟合-应用)/`run`(逐年留一)。项目唯一评估口径。 |
| `validate_methods.py` | AUC/KS/IV 指标 + `apply_woe_score`(打分应用WOE)+`make_features`。 |

### pipeline/ — (已并入 validate/)
> eval_loyo + validate_methods 已移入 `validate/`(评估底层引擎,被 train/report 共用)。pipeline/ 目录移除。

### diagnostics/ — 诊断
| 脚本 | 职责 |
|---|---|
| `diagnose_bin_drift.py` | 分箱漂移诊断。 |
| `diagnose_strategy_scorecard.py` | 策略评分卡诊断。 |

> WOE 三处(train_scorecard.woe_transform 训练 / eval_loyo.fit_woe+apply_woe CV / validate_methods.apply_woe_score 打分)是**模型训练不同阶段的同一算法**,输出格式按阶段不同,非冗余,勿统一。

---

## 常用操作

### 训练/验证从 DB 样本空间读(★新工作流,无 parquet 文件)
```python
from validate.save_validation_db import load_features
df = load_features('placement_train_20260627')   # 传 tag → DB 直读 DataFrame
df = load_features('/path/to/file.parquet')       # 传路径 → 兼容旧文件
```
训练脚本(train_scorecard_model/eval_loyo/train_to_production)已全接 `load_features`,`--features_path` 传 tag 即从 DB 取。

### 新因子 → 生产 标准流程(7 步固化,勿各搞一套)
1. **计算**:OHLCV派生→`features/factor_engine.py`加函数;tushare取数→`data/fetch_factors.py`加`ingest_X`;衍生→`features/derive_features.py`加`derive_*`族。
2. **因子验证**(必做):`validate/eval_factors.py <因子> --horizons all` 看 IV/AUC/KS×全期限+覆盖+方向。
3. **重建特征**:`features/export_features.py`→panel;`register_panel` 入 ml_train_wide。
4. **选择+训练**:`train/train_to_production.py --features_path <tag> --horizon 7 --kind gray --model sc --min-folds 12`(canonical五步自动选;**不加`--set-current`**先验证)。
5. **A/B 对真实生产**(必做,AUC+KS同报):候选共识 vs 当前生产,同 tag 同口径 LOYO;ΔAUC**和**ΔKS都报,差值才有效。
6. **评分卡审计**(上线前必做):`report/audit_scorecard.py` 导出每特征业务解释/IV/系数/分箱/分值,人审过才部署。
7. **部署**(A/B正+审计过才切):加`--set-current`;`deploy/score_one_stock.py`冒烟;`deploy/manage_models.py current`确认。回滚:`manage_models.py set full <旧版>`。

> ⚠️ **`--set-current` 只给 7m 主模型!** 1m/3m 辅助模型禁用(会冲掉 7m 主);辅助"部署"=训练即生效(predict 按 label_config 取最新)。

### 预测 / 解释
```bash
python ml_training/deploy/score_one_stock.py 300604.SZ 长川科技 20260618      # 单股终端表格
python ml_training/deploy/batch_screen_and_score.py --input <名单.xlsx>       # 批量
python ml_training/report/explain_scorecard.py 300604.SZ 长川科技 20260618    # 评分卡拆解
```

### 切换 / 查看生产模型
```bash
python ml_training/deploy/manage_models.py current
python ml_training/deploy/manage_models.py list
python ml_training/deploy/manage_models.py set full <version>   # 回滚生产
```

---

## 关键约定(必读)

1. **评估用 LOYO,不用 shuffle CV** — shuffle 同股近邻日混入 train/test 致虚高。`eval_loyo.run` 是去偏标准。
2. **DB 样本空间唯一** — 训练/验证从 `ml_train_wide` 读(`load_panel`/`load_features`),不依赖本地 parquet;新 panel `register_panel` 入库。
3. **报告全入库** — 验证→`ml_validation*`(`save_validation_run`);模型报告→`ml_model_report`(`save_model_report`);**项目内不留裸 csv**。
4. **canonical 特征选择唯一** — `features/feature_selection.py` 五步固定(IV→PSI→corr→VIF→LGBM),改阈值只改此处。
5. **三层一入口铁律** — 取数=`data/`;装配=`features/build_features`;新特征只往现有入口加函数,不新开脚本。
6. **标签防泄漏** — `features/feature_exclusions` 拦截涨跌幅/excess_/标签/报价日等;新目标列务必加排除清单。
7. **token 不硬编码** — 走 `tushare_token.resolve_tushare_token()`。
8. **模型权重在 DB**(`ml_model_meta`),版本指针在 `model_registry.json`(`current.full`)。
9. **set_current=False 扫描** — 多期限/实验模型不切生产;生产提升经 `manage_models.py set` 人工确认。
10. **评分卡精简(10–15feat)** — `train_to_production --min-folds 12`(默认3会共识膨胀→过拟合)。
11. **A/B 必对真实生产, AUC+KS 同报** — 差值才有效(绝对 AUC 不可比)。
12. **上线前必出评分卡审计** — `audit_scorecard.py` 人审,不只看 AUC。

---

## 标签设计方法论(双样本空间)

**定 lose/win 阈值用 `train/sweep_label.py` 扫描(可迁移任意期限)**。核心:训练只在非灰度样本,生产灰度筛不掉 → 必须同时评两套样本空间。

| 指标 | 样本空间 | 用途 |
|---|---|---|
| 非灰 AUC/KS | 仅清晰赢/输 | 参考;**极端阈值会虚高** |
| 灰度% | 全体 | 实战筛不掉占比 |
| **全样本 IC** | 含灰度全体 | **选阈值依据**(Spearman 概率,收益) |

**选阈值看全样本 IC,不看非灰 AUC**(极端阈值非灰 AUC 是海市蜃楼)。量纲:AUC 0.7可用/KS 0.3可用/IC月 0.02–0.05弱可用。生产口径见 `train_horizon_models.GRAY_CFG`。

---

## 关键发现与边界

- **模型本质 = regime/动量(beta),非选股 α**。超额收益标签(剥 regime)AUC 反降。
- **AUC 天花板 ~0.65–0.73**(LOYO);定增专属特征多已定价。
- **期限可预测性**:月级 0.63–0.67(中),周级 ~0.55(近噪声)。
- **有信号因子**:`chip_concentration`(2018+子集)、`smc_premium_discount`、技术/动量族。
- **无效因子**:定增结构条款、SMC 多数变体、moneyflow、超额收益标签。

---

## 依赖

- Python ≥ 3.10(用 vnpy 环境:`~/anaconda3/envs/vnpy/bin/python`)
- lightgbm, scikit-learn, pandas, numpy, pymysql, openpyxl, tushare
- MySQL(investment_valuation 库;`max_allowed_packet=1G` 供 ml_train_wide 大 panel BLOB)
