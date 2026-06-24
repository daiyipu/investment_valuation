# 数据 & 样本空间建设总纲 (Data & Sample-Space Master)

> 取数 + 特征工程 + 样本空间的**统一架构**, 跨业务类型(定增 / 回测 / 未来其他)。
> 这是 north-star: 加任何新业务前先读这里, 勿重复造取数/特征轮子。
> 关联:[`scripts/BACKTEST_PIPELINE.md`](scripts/BACKTEST_PIPELINE.md)(回测专细化) · [`ml_training/README.md`](ml_training/README.md)(训练/模型侧)。

---

## 0. 一句话

**业务类型 = 一份 manifest(样本清单)+ 一种使用方式。取数 / 特征 / 标签三层与业务类型无关, 全局共享。**

业务类型的差异**只集中在 2 个地方**:Layer 1(样本空间怎么定)+ Layer 4(算完怎么用)。中间三层(Layer 0 取数 / Layer 2 特征 / Layer 3 标签)是同一套代码、同一套 DB 表。

---

## 1. 核心抽象:`(股票, 报价日)` 原子 + 虚拟报价日

整个系统的原子 = **`(stock_code, quote_date)`**。一切下游(取数 PIT、算特征、算标签)都只认这个键。

> **虚拟报价日**:定增 vs 回测 vs 其他业务,**唯一区别是报价日怎么来**。
> - 定增:报价日 = 真实定增发行日(`issue_date`)
> - 回测:报价日 = 月末(虚拟, PIT 过滤, 排定增同时期)
> - 未来:报价日 = 事件日 / 筛选日 / 季末调仓日 / …
>
> 一旦有了 `(股票, 报价日)`, 取数 / 算特征 / 算标签的管线**完全相同**。这是系统能跨业务复用的根本。

---

## 2. 五层架构

```
            ┌── 业务类型无关(共享) ──┐  ┌── 业务类型差异点 ──┐
Layer 0  原始数据摄入 ────→ 共享 DB 表(market/fcf/financial/scores/industry/hk_hold)
Layer 1  样本空间(manifest) ★BT差异★  (股票, 报价日) 清单
Layer 2  特征工程 ─────────→ 通用核心 15 特征(compute_features) + BT 专属因子(定增结构等)
Layer 3  标签 ────────────→ forward-return(报价日, 期限)
Layer 4  使用 ★BT差异★     定增:train/score  回测:L/S  未来:rank/event-study/rotate
```

| 层 | 共享? | 业务类型差异 |
|---|---|---|
| L0 取数 | ✅ 共享(同一套 DB 表) | 仅"股票清单"不同 |
| L1 样本空间 | ❌ **差异点** | 报价日来源(真/虚拟月末/事件日) |
| L2 特征 | ✅ 核心共享(15 通用特征) | 仅 BT 专属因子(定增结构 em_*/pp_*) |
| L3 标签 | ✅ 共享(bench_return + 期限) | 仅期限选择(7m/1m/3m/…) |
| L4 使用 | ❌ **差异点** | train/score vs L/S vs rank vs … |

---

## 3. 业务类型矩阵(已实现 + 未来)

| 业务类型 | 报价日来源 | manifest | L4 使用 | 状态 |
|---|---|---|---|---|
| **定增** | 真实 `issue_date` | `placement_evaluation` | SC 评分卡 train/score | ✅ 生产 |
| **全A 回测** | 月末(虚拟, PIT, 排同时期±7月) | `backtest_samples.parquet` | L/S 组合 + IC/ICIR | ✅ pilot |
| 可转债 / 配股 / 大宗 | 事件日 | `<event>_evaluation` | 同定增 或 事件研究 | 未来 |
| 主动选股 / 量化打分 | 筛选日(任意 / 季末) | `screening_samples` | ranking 输出 | 未来 |
| 事件研究 | 事件日(披露 / 减持) | `event_samples` | 事件收益 + CAR | 未来 |
| 行业轮动 / 择时 | 周期调仓日 | `rebalance_samples` | 轮动信号 | 未来 |

**所有未来类型:Layer 0/2/3 零改动, 只需加一份 manifest + 一个使用脚本。**

---

## 4. 各层详解

### Layer 0 — 原始数据摄入(共享)
- **共享 DB 表**:`market_data`(qfq)、`historical_fcf`、`financial_indicators`(27比率)、`company_annual_scores`(总分)、`industry_data`/`industry_daily`、`hk_hold`(北向)。
- **输入 = 股票清单(唯一 BT 相关输入)**:走 `resolve_universe` 规格(placement/fullA/sample:N/file:path)或 Excel。
- **统一入口(✅ 已统一)**:[`ingest_raw.py`](scripts/ingest_raw.py) — `--universe <spec>` 或 `--src <xlsx>` → 逐股 `ingest_stock_full` → 共享表。**任意业务类型(定增/回测/未来)的 L0 都走这里。**
- **复用单股函数**:[`update_market_data.ingest_stock_full()`](scripts/data_pipeline/update_market_data.py#L1993)(market+FCF+行业, 参数化 industry_days/skip_*)。
- **定增链 [`update_all_data.py`](scripts/data_pipeline/update_all_data.py)** 是含 L0 的 12 步全链编排(指数→行情→定增名单→标签→因子→特征导出→衍生);其 L0 可选用 `ingest_raw --universe placement`(顺带补定增 batch 缺失的 FCF+行业)。回测用 `ingest_raw --universe <spec> --skip-market`(价量走运行时 pro_bar)。

### Layer 1 — 样本空间 / manifest(BT 差异点)
- **定增**:`placement_evaluation`(真报价日)。
- **回测**:[`gen_backtest_samples.py`](scripts/data_pipeline/gen_backtest_samples.py) 产 `backtest_samples.parquet`(股票×月末虚拟报价日, PIT 过滤在市/非ST/非次新/退市前纳入, **排除定增同时期±7月**避免标签窗口重叠泄漏)。
- **未来**:各 BT 的 manifest 生成器(报价日语义不同, 结构同)。
- **理想**:`gen_backtest_samples` 泛化为 `gen_samples --mode placement/backtest/event/...`。

### Layer 2 — 特征工程(核心共享 + BT 专属)
- **通用核心(15 特征, 任何 BT 可算)** — `compute_features(stock, date, feat_list)`:
  - 价量:alpha/beta(运行时 pro_bar qfq 全量预取 + `≤date` 切片)
  - 行业相对:行业 PE/PB 增长(`industry_daily`)
  - 财务:27 比率(`financial_indicators`, ann_date≤date)
  - 5 特殊:FCF_加速 / 总分_delta_2y / nb_hold / PB_vs同行 / sue_beat
  - **PIT 铁律**:所有特征 `ann_date/year/trade_date ≤ 报价日`。
- **BT 专属因子(定增)**:`fetch_factors` 族的 `em_*/pp_*`(定增结构)、`chip_*`、`mf_*/nb_*`、`smc_*`。这些绑定 placement_evaluation 结构, 不 portable。
- **5 特殊特征等价性(2026-06 核实)**:[`feature_loaders.py`](scripts/feature_loaders.py) 是任意 BT 的 PIT 规范 loader;定增侧历史另有实现(批量快照回写 pe 再读)。逐特征:
  - `FCF_加速` / `总分_delta_2y` / `nb_hold_ratio`:**数学等价**(pit_max 同、tail(3)、iloc[-1]), 可安全统一。
  - `sue_beat`:**共享同一函数**(`fetch_factors._sue_for_sample`), 已统一。
  - `PB_vs_同行中位`:**❌ 口径冲突** —— 定增 = 个股PB / **peer_companies 同行中位**(非 PIT 快照, 覆盖53%); feature_loaders = daily_basic PB / **industry_daily 行业PB**(PIT ≤date, 覆盖91%)。生产 SC 模型 15 特征含此列 → 回测喂行业 PB 与训练(peer)分布不同(中位 1.17 vs 1.70)→ **PB 特征系统性错配, 回测 IC/L-S 被部分污染**。统一唯一卡点, 须先定口径(建议采 PIT 行业口径 → 重训定增模型, 顺带修定增 PB 的 PIT 泄漏)。

### Layer 3 — 标签(共享, 报价日+期限参数化)
- **引擎**:[`compute_labels.bench_return`](ml_training/compute_labels.py)/ `add_months` / `_nearest_close`(forward-return, 吸附最近交易日)。
- **定增**:`recompute_label_qfq.py`(return_*m 写回 placement_evaluation)。
- **回测**:`fwd_returns(报价日, months=7)` 写入 panel。
- **口径**:7m = 定增解禁期 = 模型标签口径(同口径才能比)。1m/3m 短期灰度列同机制。
- **未来**:同一 forward-return 引擎, 换期限即可。

### Layer 4 — 使用(BT 差异点)
- **定增**:`train_scorecard` / `score_sc` → WOE+LR 评分卡 → 档位(训练集十分位校准)。
- **回测**:[`backtest_long_short.py`](scripts/backtest_long_short.py) 读 panel → `score_sc` 打分 → Top/Bottom 多空 + 月度 IC/ICIR + 12 日历月分组夏普。
- **未来**:rank 输出 / 事件 CAR / 轮动信号 —— 各写自己的薄使用层。

---

## 5. 新增业务类型:3 步接入法

```
① 写 manifest 生成器(报价日语义)→ (股票, 报价日) 清单
② 【零改动】复用 Layer 0(摄入, 若股票已在 universe)/ Layer 2(compute_features)/ Layer 3(bench_return)
③ 写 Layer 4 使用脚本(读 features_<bt>.parquet → 该 BT 的输出)
```

**示例 — 可转债**:
1. `gen_samples --mode cb` → 从可转债事件表取发行日作报价日 → `cb_samples.parquet`。
2. 股票大多已在全A universe → 表已有数据;`build_panel --src cb_samples --out features_cb.parquet`(复用 compute_features)。
3. 用 `score_sc` 打分(若可转债 regime 接近定增)或写事件收益脚本。

**唯一可能要补数据的情况**:新 BT 引入**新股票**(不在现有 universe)→ 跑 `ingest_stock_full` 补该股原始数据。新 BT 引入**新特征**(如可转债条款因子)→ 在 Layer 2 加 loader, 但通用 15 特征不变。

---

## 6. 历史数据 vs 增量数据(全局)

**历史(一次性全量回填, 任何 BT 受益)**:
- 全A × 2010-2025 原始数据落共享表(`ingest_stock_full` industry_days=4000 + `backfill --force --years 16` + `batch_score --years-wide 16`)。
- 做完一次, 表常驻; 任何 BT(定增/回测/未来)直接读, 不重取。

**增量(持续)**:
- 新月/新事件 → 重跑 manifest(加新报价日)→ panel 只补新截面。
- 新财报 / 新行业日线 → 摄入脚本幂等 upsert(DELETE+INSERT / ON DUPLICATE KEY), 只追加新数据。
- 新股 → `fetch_universe --refresh` → `ingest_stock_full` 补。
- 新披露(SUE)→ `prefetch_sue_timelines(refresh=True)`(历史 PIT 固定, 平时不刷)。
- **panel 是缓存产物**:原始数据/特征清单不变就直接读(秒级)。

---

## 7. 现状诚实评估 + 待统一(2 处)

| 处 | 现状 | 理想 | 优先级 |
|---|---|---|---|
| Layer 0 摄入 | ✅ **已统一** `ingest_raw.py`(resolve_universe → ingest_stock_full, 任意 BT) | — | 已完成 |
| Layer 2 specials | 5 特征两套实现, 但 4/5 数学等价(FCF/总分/nb_hold/sue_beat, sue_beat 已共享函数); **PB_vs_同行中位 口径冲突**(定增 peer 中位 vs 回测行业 PB, 且是模型特征 → 回测 PB 错配) | 定 PB 口径(建议采 PIT 行业)→ 重训定增模型 → 全部并 feature_loaders | P2(待决策) |
| Layer 1 manifest | gen_backtest_samples 回测专用 | 泛化 `gen_samples --mode <bt>` | P3 |

**结论**:L0 已统一(ingest_raw 是任意 BT 的唯一 L0 入口);剩 2 处"双套"(L2 specials / L1 manifest)是历史补丁, 不阻塞新 BT 接入(新 BT 照样能复用), 统一后更省维护。

---

## 8. 全仓文件地图(按层)

| 层 | 文件 | 职责 |
|---|---|---|
| **L0 取数** | `ingest_raw.py` | **统一 L0 入口**(resolve_universe → ingest_stock_full, 任意 BT) |
| | `refresh_industry_daily.py` | 按唯一行业(去重)刷 industry_daily **系列**全历史(避免逐股撞 sw_daily 限频; 不动映射) |
| | `refresh_industry_mapping.py` | 补 industry_data **股票→行业映射**(只分类 `index_member_all`, 不动系列; `--src features.parquet` 补训练集) |
| | `data_pipeline/update_market_data.py` | `ingest_stock_full`(单股 market+FCF+行业) + `ingest_raw` 复用 |
| | `data_pipeline/update_all_data.py` | 定增 12 步全链编排(含 L0) |
| | `data_pipeline/backfill_financial_indicators.py` | 27 财务比率(三表算法, `--force`) |
| | `batch_financial_score.py` | 年度总分(EFAES, `--years-wide`) |
| | `data_pipeline/fetch_factors.py` | 定增结构/筹码/资金流/SMC + SUE 时间线 |
| **L1 样本** | `data_pipeline/fetch_universe.py` | 全A清单 + PIT 元数据 + `resolve_universe()` |
| | `data_pipeline/gen_backtest_samples.py` | 回测 manifest(月末虚拟报价日 + 排同时期) |
| **L2 特征** | `ml_training/derive_features.py` | alpha/beta + 行业相对估值衍生 |
| | `ml_training/export_features.py` | `load_financial_ratios` / `load_scored`(定增锚定) |
| | `feature_loaders.py` | 5 特殊特征 PIT loader(回测) + SUE 落盘 |
| **L3 标签** | `ml_training/compute_labels.py` | `bench_return`/`add_months`(forward-return) |
| | `data_pipeline/recompute_label_qfq.py` | 定增 return_*m 写回 |
| **L4 使用** | `ml_training/train_to_production.py` | 定增:端到端训练→生产 |
| | `ml_training/predict_profitability.py` | `score_sc`(SC 打分, 共享) |
| | `backtest_long_short.py` | 回测:L/S + IC |
| | `build_backtest_panel.py` | panel 构建(compute_features + fwd_returns) |

---

## 9. 约定(全局)
- 环境:**vnpy py3.10**(`/Users/davy/anaconda3/envs/vnpy/bin/python`), cwd 显式 `cd price_maintenance_risk_analysis/`。
- token:走 `resolve_tushare_token()`, 绝不硬编码。
- PIT:所有取数 `≤报价日`;标签 forward `>报价日`。
- panel = parquet 缓存(列式, 批量读);DB = 源数据(PIT 查询)。
- 新 BT 先确认股票在 universe(否则先摄入), 再复用 compute_features, 勿重写取数。
