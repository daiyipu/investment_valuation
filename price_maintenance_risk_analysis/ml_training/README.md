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

### 新因子 → 生产 标准流程(固化, 每次开发新指标按此走, 勿各搞一套)

> 教训(2026-06-20): 跳过验证报告 / A/B 不对真生产 / 不出审计, 凭单因子 IV 就部署 → 74feat 过拟合评分卡、假 +AUC、KS 丢失。下述 7 步**必走**, 顺序不可省。

**Step 0 计算** — 按数据源三选一:
- OHLCV 派生(量价/技术/动量/Beta)→ `factor_engine.py` 加函数 + 注册 `_FACTOR_NAMES_OF`; 经 `derive_alpha_beta_factors` 自动进 features_derived(免发射)。
- tushare 取数(基本面/事件/资金)→ `scripts/data_pipeline/fetch_factors.py` 加 `ingest_X`+`COLS`+`SOURCES` → 写 `placement_evaluation` DB。
- 衍生算术 → `derive_features.py` 加 `derive_X` 族。
- 约定: token 走 `resolve_tushare_token()`; 每股1次全历史调用 + 3重试 + 股间 sleep; **单位注意**(如 forecast 净利万元×1e4); 命名避开 `feature_exclusions` 子串; 发比率/z-score 不发绝对元。

**Step 1 发射**(仅 DB 源需要): `export_features.py` 加前缀循环(仿 `smc_*`/`sue_*`)或 `pe.get`; derive 源免改。

**Step 2 因子验证报告(必做, 发给审核)** — `eval_factors.py`:
```bash
python ml_training/eval_factors.py 新因子 --horizons all                                        # DB 源
python ml_training/eval_factors.py 新因子 --parquet data/features_derived.parquet --horizons all # derive 源
```
看 IV/AUC/KS × 全期限 + 覆盖 + 方向。**关键判断**: 信号在哪个期限? 7m 弱 → 路由 1m/3m 辅助模型, 不强塞 7m 主模型。单因子 IV(尤其错期限)≠评分卡价值。

**Step 3 重建特征**: `export_features.py`→features.parquet; `derive_features.py`→features_derived.parquet。

**Step 4 选择 + 训练(精简评分卡)** — `train_to_production.py`:
```bash
python ml_training/train_to_production.py ml_training/data/features_derived.parquet \
       --horizon 7 --kind gray --model sc --min-folds 12
```
- canonical 五步自动选(`feature_selection.py`: **IV>0.02 且 top50** → PSI≤0.25 → |r|≤0.7 → VIF<5 → LGBM)。
- **`--min-folds 12`** 保共识 10–15feat(默认 3 会膨胀到几十→过拟合; 评分卡要精简可解释)。
- `loyo_fixed` 出真 LOYO AUC+KS(LOYO 不用 shuffle CV)。**不加 `--set-current`**(先验证再切)。

**Step 5 A/B 对真实生产(必做, AUC+KS 同报)**: 候选共识 vs 当前生产(锁定特征集), 同 parquet 同口径 LOYO。
- ΔAUC **和** ΔKS 都报; 须超噪声底(~+0.5pp)且 KS 不降才考虑上。
- 绝对 AUC 不可比(一次性选择不稳), **差值才有效**。

**Step 6 评分卡审计(上线前人审, 必做)** — `audit_scorecard.py`:
```bash
python ml_training/audit_scorecard.py            # 导出 current.full 评分卡审计
```
每特征: 业务解释+IV+LR系数+贡献度+分箱+各箱WOE+分值 → `output/audit_<ver>/scorecard_audit.{csv,md}`。审核逐箱核, 不只看 AUC。业务解释维护在 `feature_glossary.py`(`feature_glossary.explain()`, 新特征补 EXPLICIT)。

**Step 7 部署(A/B 正 + 审计过 才切)**:
```bash
python ml_training/train_to_production.py ml_training/data/features_derived.parquet \
       --horizon 7 --kind gray --model sc --min-folds 12 --set-current
python scripts/score_one_stock.py 300604.SZ 长川科技 20260618   # 冒烟确认端到端
python ml_training/manage_models.py current                     # 确认新版本已切
```
回滚: `manage_models.py set full <旧版本>`。

> ⚠️ **主 vs 辅助模型 — `--set-current` 只给 7m 主模型用!**
> registry 只有 `current.full`/`current.scorecard` 两个指针(无按期限分)。`--set-current` 写 `current.full` → **仅 7m 主模型**用。
> **1m/3m 辅助模型禁用 `--set-current`**(会冲掉 7m 主)! 辅助训练:
> ```bash
> python ml_training/train_to_production.py ml_training/data/features_derived.parquet \
>        --horizon 1 --kind gray --model sc --min-folds 12     # 不加 --set-current
> # 同理 --horizon 3。predict 按 label_config('1m_gray_sc'/'3m_gray_sc') 取最新版自动生效。
> ```
> predict 多期限补充列(`predict_profitability.py` 9c 段)按 label_config 取最新 SC, 与 current.full 无关 → 辅助"部署"=训练即生效(成最新版), 无需 set。

### 评估因子效力(Step 2 工具)
```bash
python ml_training/eval_factors.py smc_premium_discount chip_concentration   # DB 指定
python ml_training/eval_factors.py --prefix smc_                              # DB 全 smc_*
python ml_training/eval_factors.py <因子> --parquet data/features_derived.parquet --horizons all  # derive 因子
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
4. **canonical 特征选择唯一** — `feature_selection.py` 五步固定(**IV>0.02 且 top50** → PSI≤0.25 → |r|≤0.7 → VIF<5 → LGBM),不再各脚本各搞一套。改阈值只改此处常量。
5. **标签防泄漏** — `feature_exclusions` 拦截 涨跌幅/周涨跌幅/excess_/标签/报价日 等。新增目标变量列务必加进排除清单。
6. **token 不硬编码** — 走 `tushare_token.resolve_tushare_token()`(旧 literal 已泄漏需轮换)。
7. **模型权重在 DB**(`ml_model_meta`),版本指针在 `model_registry.json`(`current.full`)。
8. **set_current=False 扫描** — 多期限/实验模型一律不切生产,避免静默切换 regime;生产提升经 `manage_models.py set` 人工确认。
9. **评分卡要精简(10–15feat)** — `train_to_production --min-folds 12`(默认 3 会共识膨胀到几十→过拟合; 2026-06-20 教训: 74feat 评分卡比 11feat 生产差)。
10. **A/B 必对真实生产, AUC+KS 同报** — 候选 vs 当前生产锁定特征集, 同 parquet 同口径 LOYO; ΔAUC **和** ΔKS 都报, 差值才有效(绝对 AUC 不可比)。只报 AUC 丢 KS = 失职。
11. **上线前必出评分卡审计** — `audit_scorecard.py` 导出每特征业务解释/IV/贡献度/分箱/分值(CSV+MD), 人审过才部署; 不只看 AUC。
12. **单因子 IV ≠ 评分卡价值(期限匹配)** — 选特征按目标期限(7m)IV; 信号在 1m/3m 的因子路由辅助模型, 不强塞 7m 主模型(会挤掉动量反变差)。

---

## 关键发现与边界(2026-06 因子挖掘结论)

- **模型本质 = regime/动量(beta),非选股 α**。超额收益标签(剥 regime)实测 AUC 反降 → 现有动量特征赚的就是 regime 的钱。
- **AUC 天花板 ~0.65–0.73**(LOYO)。定增专属特征多已定价;非公开筹码类(chip)有信号但受历史覆盖陷阱。
- **期限可预测性**:月级 0.63–0.67(中),**周级 ~0.55(近噪声)**。短期收益=噪声主导。
- **聪明钱(SMC)**:仅 `premium_discount`(区间位置, 反向均值回归)微弱有效(IV 0.035–0.054);ote/bos/fvg/displacement/sweep ≈0。周/月线 bar 配月标签是关键(日线配月标签错配)。
- **有信号的因子**:`chip_concentration`(2018+ 子集 +2.5pp, 覆盖陷阱)、`smc_premium_discount`、技术/动量族。
- **无效因子**:定增结构条款(公开已定价)、SMC 多数变体、moneyflow(94%覆盖仍~0)、超额收益标签。

---

## 标签设计方法论(双样本空间, 2026-06-20 固化)

**定标签 lose/win 阈值不靠拍脑袋, 用 `sweep_label.py` 扫描; 可迁移任意期限(1/3/6/7/12m)。**

核心:**训练只在非灰度样本上(灰度被丢), 但生产时灰度筛不掉 → 必须同时评两套样本空间**, 否则会被极端阈值的"高 AUC"误导。

```bash
python ml_training/sweep_label.py ml_training/data/features_derived.parquet --horizon 3
# 固定 win(+10, 生产口径), 扫 lose [0,-5,-10,-15,-20], LOYO 报:
```

| 指标 | 样本空间 | 含义 | 用途 |
|---|---|---|---|
| **非灰 AUC/KS** | 仅清晰赢/输 | 训练口径区分度 | 参考; **极端 lose 阈值会虚高**(只留极值) |
| **灰度%** | 全体 | 实战筛不掉的灰度占比 | 越高→有效样本越少 |
| **全样本 IC** | 含灰度全体 | Spearman(预测概率, 实际收益) 实战排序 | **选阈值的依据** |

**选阈值看【全样本 IC】, 不看非灰 AUC。** 教训(1m):
- `(−20,10)` 非灰 AUC 0.61 看着最强, 但灰度 75%、全样本 IC≈0 → **海市蜃楼**(尾部区分不等于全人群排序)。
- 真实甜点 `(−5,10)/(−10,10)`:全样本 IC 最高(~0.039), 灰度 45–61%。

**指标量纲**(选阈值/判好坏用):
- **AUC**(全阈值排序):0.5 随机 / 0.7 可用 / 0.8 强。
- **KS**(最佳单阈值 TPR−FPR 差):0.2 中 / 0.3 可用 / 0.4 强。AUC 低+KS 中 = 有一个切分点但整体排序弱(尾部型)。
- **IC**(rank corr, 月级):<0.02 噪声 / 0.02–0.05 弱但可用 / 0.05–0.10 中 / >0.10 强。
- 生产口径见 `train_horizon_models.GRAY_CFG`(改阈值 = 改此处 + 重跑本扫描确认)。

---

## 依赖

- Python ≥ 3.10(用 vnpy 环境:`~/anaconda3/envs/vnpy/bin/python`, 非 base py3.7)
- lightgbm, scikit-learn, pandas, numpy, pymysql, openpyxl, tushare
- MySQL(investment_valuation 库:placement_evaluation/market_data/relative_valuation/historical_fcf/industry_daily/industry_data/issue_date_locked/ml_model_meta/ml_dataset_snapshot 等)
