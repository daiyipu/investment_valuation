# 全A 多空回测链路(Backtest Pipeline)

> 定增 SC 评分卡(15特征, 7m horizon)在全A上的组合回测:**信号能不能赚钱**。
> 本文档 = 回测数据/样本空间的**唯一入口**, 建新回测前先读这里, 勿重复造轮子。
> 上位文档:[`../DATA_PIPELINE.md`](../DATA_PIPELINE.md)(取数+特征+样本空间**总纲**, 跨业务类型:定增/回测/未来)。

---

## 一、一图看懂(5 步)

```
①universe  →  ②样本清单  →  ③原始数据摄入  →  ④特征panel  →  ⑤回测
 全A清单       股票×月末       落共享DB表        算15特征+7m收益   读panel打分多空
               PIT过滤          (定增也读这套)     → parquet缓存    IC/ICIR/CAGR
               排同时期
```

**核心思想 = 虚拟报价日**:定增 vs 全A 唯一区别是报价日真假。取数/算特征只认 `(股票, 回溯日期)`。
回测把**月末**当虚拟报价日, 现有取数管线照跑, 不改结构。

---

## 二、标准命令序列(从 0 到回测)

环境:**vnpy py3.10**, cwd 必须 `price_maintenance_risk_analysis/`。

```bash
PKG=price_maintenance_risk_analysis
PY=/Users/davy/anaconda3/envs/vnpy/bin/python
cd /Users/davy/github/investment_valuation

# ① universe(全A ~5000; 首次或定期刷新)
$PY $PKG/scripts/data_pipeline/fetch_universe.py                         # → data/universe.parquet
$PY $PKG/scripts/data_pipeline/fetch_universe.py --sample 500           # pilot: 额外产分层抽样

# ② 样本清单(股票 × 月末虚拟报价日, PIT 过滤, 排定增同时期 ±7月)
$PY $PKG/scripts/data_pipeline/gen_backtest_samples.py --universe sample:500 --years 2010-2025   # pilot
$PY $PKG/scripts/data_pipeline/gen_backtest_samples.py --universe fullA    --years 2010-2025     # 全量

# ③ 原始数据摄入(全历史; 统一 L0 driver + 两个财务脚本吃 backtest_universe.xlsx)
$PY $PKG/scripts/ingest_raw.py --universe sample:500 --skip-market        # FCF+行业(回测价量走运行时 pro_bar; industry_days=4000→start≈2004)
$PY $PKG/scripts/data_pipeline/backfill_financial_indicators.py $PKG/ml_training/data/backtest_universe.xlsx --force --start-year 2009 --end-year 2025 --years 16
$PY $PKG/scripts/batch_financial_score.py $PKG/ml_training/data/backtest_universe.xlsx --years-wide 16

# ④ 特征 panel(复用 compute_features; 首次跑, 之后复用缓存)
$PY $PKG/scripts/build_backtest_panel.py --horizon 7                      # → data/features_backtest.parquet

# ⑤ 回测(读 panel 打分多空; 秒级)
$PY $PKG/scripts/backtest_long_short.py --horizon 7                       # IC/ICIR + L-S CAGR/maxDD
```

**决策门**:pilot `ICIR>0.3 且 L-S 年化为正 且 12日历月组夏普多为正` → 信号未严重衰减 → 扩全A; 否则报告衰减。

---

## 三、文件作用清单(谁读谁写, 何时重建)

### 回测专属脚本(本链路新建)

| 文件 | 职责 | 输入 | 输出 | 何时重跑 |
|---|---|---|---|---|
| `data_pipeline/fetch_universe.py` | 全A清单 + ST/次新/退市 PIT 元数据 + 分层抽样 | tushare stock_basic | `data/universe.parquet`, `data/namechange.parquet`, `data/universe_sample_N.parquet` | universe 变化(新股/退市)时 `--refresh` |
| `data_pipeline/gen_backtest_samples.py` | 样本清单(股票×月末) PIT 过滤 + 排定增同时期 | universe.parquet, placement_evaluation | `data/backtest_samples.parquet`, `data/backtest_universe.xlsx` | 改 universe/年份/排除窗口时 |
| `ingest_raw.py` | **统一 L0 摄入**(任意业务类型, 复用 `ingest_stock_full`) | `--universe <spec>` 或 `--src <xlsx>` | DB: `market_data`, `historical_fcf`, `industry_data`, `industry_daily` | 新股/补历史(`--industry-only` 只补行业) |
| `build_backtest_panel.py` | 逐截面算 15 特征 + 7m 收益 → panel | backtest_samples.parquet + DB表 + SUE缓存 | `data/features_backtest.parquet` | 样本/原始数据/模型特征清单变时 |
| `backtest_long_short.py` | 读 panel → `score_sc` 打分 → 多空/IC | features_backtest.parquet | stdout(IC/ICIR/CAGR/maxDD) | 每次回测(秒级, 读缓存) |
| `feature_loaders.py` | 5特殊特征 PIT loader(FCF/总分/nb_hold/PB_vs/SUE) | DB + tushare API | DataFrame(内存) + `data/sue_timelines.parquet` | 被 build_backtest_panel 调用; SUE落盘后免重取 |

### 复用脚本(定增/回测共用, **勿改其结构**)

| 文件 | 复用点 | 回测怎么喂 |
|---|---|---|
| `data_pipeline/update_market_data.py` | `ingest_stock_full()`(单股摄入 market/FCF/行业) | `ingest_raw` 批量调, 回测传 `--skip-market`(价量走运行时 pro_bar) |
| `data_pipeline/backfill_financial_indicators.py` | 三表算法 27 比率(与定增同源) | 喂 backtest_universe.xlsx + `--force`(新股建行) |
| `batch_financial_score.py` | EFAES 年度评分 | 喂 backtest_universe.xlsx + `--years-wide 16`(全历史) |
| `data_pipeline/fetch_factors.py` | `_build_disclosure_timeline` / `_sue_for_sample`(SUE) | feature_loaders 调 |
| `ml_training/derive_features.py` | `derive_alpha_beta_factors` / `derive_industry_valuation_growth` | backtest_long_short.compute_features 调 |
| `ml_training/export_features.py` | `load_financial_ratios`(PIT 财务) | compute_features 调(**注**:`load_scored` 锚定 placement, 回测不用, 故有 feature_loaders) |
| `ml_training/predict_profitability.py` | `score_sc`(SC 打分) | 回测/定增共用 |

### DB 共享表(定增 + 回测都读写, **集成点**)

| 表 | 谁写 | 谁读 | 回测特征 |
|---|---|---|---|
| `market_data` | update_market_data | (回测不读, 走运行时 pro_bar) | — |
| `historical_fcf` | ingest_stock_full | feature_loaders | FCF_加速 |
| `financial_indicators` | backfill_financial_indicators / batch_financial_score | export_features.load_financial_ratios | ar_turn/debt_to_eqt 等 |
| `company_annual_scores` | batch_financial_score | feature_loaders | 总分_delta_2y |
| `industry_data` / `industry_daily` | ingest_stock_full | derive_industry / feature_loaders | 行业PE_250d / PB_vs同行 |
| `hk_hold` | (tushare pro.hk_hold 运行时) | feature_loaders | nb_hold_ratio |
| `placement_evaluation` | fetch_placements | gen_backtest_samples(排同时期) | — |

### 关键产物(parquet 缓存)

| 产物 | 含义 | 何时失效重建 |
|---|---|---|
| `data/universe.parquet` | 全A清单 + list/delist_date/name | `--refresh` |
| `data/backtest_samples.parquet` | 回测样本脊柱(股票×月末) | universe/年份变 |
| `data/backtest_universe.xlsx` | 唯一股(喂摄入脚本) | 随 samples 重建 |
| `data/features_backtest.parquet` | **回测panel**: N行 × 15特征 + return_7m | 样本/原始数据/特征清单变 |
| `data/sue_timelines.parquet` | SUE披露时间线全历史(含空股占位) | 增量新披露 → `refresh=True` |

---

## 四、历史数据 vs 增量数据

**历史数据(一次性全量回填)**:
- 机制现成: `ingest_stock_full`(industry_days=4000→start≈2004) + `backfill --force --years 16` + `batch_score --years-wide 16`。
- 做完一次, 表里就是 **全A × 2010-2025**, 后续回测直接读, 不再取数。
- 成本: 数小时 tushare(5000 积分无限频; sw_daily 5次/天→降级 AKShare)。

**增量数据(持续)**:
- **新月份**: 重跑 gen_backtest_samples(加新月末)→ panel 只需算新截面(或全重建, 旧截面缓存复用)。
- **新财报/新行业日线**: 摄入脚本幂等 upsert(DELETE+INSERT / ON DUPLICATE KEY), 只追加新数据。
- **新股**: `fetch_universe --refresh` → 摄入新股票。
- **新披露(SUE)**: `prefetch_sue_timelines(codes, refresh=True)` 重取覆盖(历史披露 PIT 固定, 不需刷新)。
- panel 是**缓存产物**: 原始数据/特征清单不变就直接读(~秒级)。

---

## 五、本次新增/修正的参数(2026-06)

| 参数 | 文件 | 作用 | 必用场景 |
|---|---|---|---|
| `--force` | backfill_financial_indicators | 全量重算每只(新股建行) | 回测全A(不加则 null-filter 漏掉无行新股→财务特征仅~50%) |
| `industry_days=4000` | ingest_raw(--industry-days) / ingest_stock_full | 行业日线回溯→start≈2004 | 2010+回测(days=3000→start≈2010-01 不够 250d 回看) |
| `--years-wide N` | batch_financial_score | score 窗口宽度(默认5, 回测16) | 早期年份总分覆盖(默认5年只覆盖近期) |
| `sue_timelines.parquet` | feature_loaders | SUE 时间线落盘 | panel 重建免重取(500股≈15min→秒级) |

---

## 六、集成度(诚实)

- **数据层 + 底层特征函数 = 已集成**(共享 DB 表 + 共享 derive/load/score_sc)。
- **摄入编排 = 已统一**(`ingest_raw.py` 是 L0 唯一入口, 任意业务类型; 定增 `update_all_data` 是含 L0 的全链编排, 可选用 `ingest_raw --universe placement`)。
- **5特殊特征 = 两套 loader**(定增 `load_scored` 锚定 placement 不可复用; 回测 `feature_loaders` 独立 PIT loader)。**这是最大的"补丁"**, P2 待统一。

---

## 七、排错速查

| 现象 | 根因 | 处理 |
|---|---|---|
| 财务特征覆盖率~50% | 未加 `--force` | 重跑 backfill `--force --years 16` |
| 行业PE_250d 早期缺失 | industry_days 太小 | `--industry-days 4000` 重摄行业 |
| 总分早期缺失 | score 窗口5年 | `--years-wide 16` |
| panel 重建慢(SUE) | 时间线未落盘 | 确认 `data/sue_timelines.parquet` 存在; 首次跑会建 |
| 78% 行业匹配 | 旧 industry_days=500 | 已修(4000); 重摄即可 |
| cwd 报错 / py3.7 | base 环境 | 用 vnpy py3.10, 显式 cd PKG |
