# DB Schema 说明书（investment_valuation 库）

> **一处维护全特征的数据地基。** 每张表:存什么 / **时序 vs 快照** / 主键 / 谁写 / 谁读 / PIT 能力 / 覆盖。
> **特征层只读"时序表"**(可 PIT 切片 ≤回溯日期);**快照表已标退役**(非 PIT、有泄漏风险,仅报告/参考用,不再喂特征)。
> 新增数据源 → 落**已有时序表**或建**新时序列**(trade_date/year/ann_date 键),**绝不新建快照表**。

## 时序表（特征层唯一数据源,可 PIT 切片）✅

| 表 | 存什么 | 主键 | 写入脚本(段2 data/) | 读取(loader/derive) | PIT 切法 | 覆盖 |
|---|---|---|---|---|---|---|
| `stock_qfq_daily` | 个股前复权日线 OHLCV | (stock_code, trade_date) | `data/market.backfill_market_data` → derive `_qfq_save` | `_OHLCV_CACHE`(derive 全族行情) | `searchsorted ≤date` | 5603股 × 全历史 |
| `stock_daily_basic` | 个股日线 daily_basic(turnover/volume_ratio/**pb**/pe/ps 待补) | (stock_code, trade_date) | `data/market` → derive `_db_basic_save` | `_DAILY_BASIC_CACHE` + `load_pb_vs_industry`(个股PB) | `searchsorted ≤date` | ~97%股 |
| `stock_monthly_hfq` | 个股后复权月线 close | (stock_code, trade_date) | `data/market` → derive `_monthly_save` | `_MONTHLY_CACHE`(derive_monthly_trend) | `searchsorted ≤date` | ~95%股 |
| `industry_daily` | 申万行业指数日线 pe/pb | (index_code, trade_date, data_source) | `data/industry.refresh_industry_daily` + ingest_stock_full | `load_pb_vs_industry`(行业PB) + derive_industry_valuation_growth | `searchsorted ≤date` | 308行业 × 全历史 |
| `index_factor_pro` | **指数技术面 87 因子**(MACD/KDJ/RSI/BOLL/MA/EMA/CCI/DMI/OBV/MFI/ROC/TRIX/WR/VR/ATR…, 大盘+申万行业) | (index_code, trade_date) | `data/industry.refresh_industry_daily`(ingest_idx_factor_pro) | 特征层 PIT 切片(≤回溯日期)→ 指数技术回溯特征 | `searchsorted ≤date` | 大盘+370行业 × 全历史 |
| `historical_fcf` | 个股年报 FCF 全历史(营收/NOPAT/FCF/营业利润/净利润_fcf/资本支出/折旧/营运资金变动 × year) | (stock_code, year) | `data/financial` + ingest_stock_full | `_FCF_CACHE` → `load_fcf_bulk`/`load_fcf_accel` | `year ≤ pit_year(date)` | ~5400股 × 16年 |
| `company_annual_scores` | 个股年度财务评分(总分/盈利能力/成长能力 × report_year) | (stock_code, report_year) | `data/financial.batch_financial_score` + `bulk_backfill_early_scores` | `_SCORE_CACHE` → `load_total_score_delta2y`/`_bulk_score_base` | `report_year ≤ pit_year(date)` | ~5400股 |
| `financial_indicators` | 个股年报 27 财务比率(流动/速动/周转/ROA/ROE/净利率/资产负债/…)× ann_date | (stock_code, report_year) | `data/financial.backfill_financial_indicators` | `load_financial_ratios` | `ann_date ≤ date`(退化 report_year≤年) | ~5360股 |
| `placement_evaluation` | **定增主表 + 多期限收益标签**(stock_code × issue_date;return_1w/2w/4w/1m/3m/7m + chip/mf/smc/nb/em/pp 列) | (stock_code, issue_date) | `data/factors.fetch_factors`(UPDATE 各列) + `compute_labels` + `backfill_evaluations` | 训练脊柱(定增);标签源 | 键值(issue_date=真报价日) | ~6400定增 |

## 键值/标签表

| 表 | 存什么 | 主键 | 用途 | PIT |
|---|---|---|---|---|
| `placement_evaluation` | (见上, 既是时序标签也是定增因子载体) | (stock_code, issue_date) | 定增训练脊柱 + 多期限标签 + chip/mf/sue 等因子回填 | 键值 |
| `issue_date_locked` | 报价日价格/MA20(定增) | (stock_code, issue_date) | 报价日精确取价 | 键值 |

## 快照表（⚠️ 非 PIT, 特征层退役中 — 仅报告/参考用）🔴

> 这些表**每股一行**(无 trade_date 维度),读取即可能引入报价日后的信息(泄漏)。特征层不再读;**统一后改走时序表 PIT 重算**。

| 表 | 存什么 | 问题 | 退役去向(改读哪) |
|---|---|---|---|
| `market_data` | 个股行情统计(MA/波动率/年化收益/胜率/当前价/中位价/价格标准差/漂移率)+ price_series JSON | 每股一行=最新一次定增快照,非PIT | → `derive_market_stats_from_ohlcv`(从 stock_qfq_daily PIT 重算) |
| `relative_valuation` | 个股/行业 PE/PB/PS 一行 | 快照,有未来信息 | → 新增 PE/PS PIT loader(stock_daily_basic + industry_daily,PIT) |
| `peer_companies` | 同行 PE/PB/PS/市值 列表 | 快照聚合,非PIT(pb-calizer 泄漏根因) | → `load_pb_vs_industry`(行业 PIT)+ 同行 PE/PS PIT 聚合 |
| `industry_data` | 每股行业映射 + 报价日行业统计 | 行业映射(静态,可留);统计快照(退役) | 映射留;统计 → `industry_daily` PIT |
| `placement_params` | 定增参数(溢价率/锁定期/无风险利率/Beta/…) | 每股一行,定增专属 | 定增结构派生用(skip_placement 时回测跳过) |
| `screening_results` | 筛选步骤结果 | 快照 | 报告用,不喂特征 |
| `stocks` | 股票基础(代码/名称/行业代码) | 静态基础 | 留(基础映射) |
| `market_indices` | 指数按 locked_date 统计 | 快照 | → `derive_market_index_features`(运行时 index_daily PIT) |
| `em_industry_boards` / `em_industry_stocks` | 东财行业板块/成分 | 静态 | universe 辅助 |

## ML 元数据表

| 表 | 存什么 | 主键 | 谁写 | 谁读 |
|---|---|---|---|---|
| `ml_model_meta` | 模型版本一行/版本(version/label_config/kind/horizon/features JSON/metrics/lr_bundle BLOB/…) | version | `db_model_store.save_model_meta` | predict/load_predict_bundle + 特征护栏 `_load_known_features`(取 features 并集) |
| `ml_dataset_snapshot` | 训练集 parquet 快照(sha256 去重) | (version, hash) | `db_dataset_store` | 复现训练集 |
| `ml_validation*`(4表) | 回测验证结果 | — | `save_validation_db` | 报告 |

## DDL 归属（待 Phase 3 整理）
- **快照表 DDL** 在 `utils/db_manager.py`(数据层,正确)。
- **时序表 DDL(stock_qfq_daily/daily_basic/monthly)** 当前错位在 `ml_training/derive_features.py`(`_qfq_save`/`_db_basic_save`/`_monthly_save` 幂等建表)→ Phase 3 迁回 `data/`(数据层建表,ML 层只读)。

## 新增数据源的铁律
1. 优先落**已有时序表**(加列:如 daily_basic 补 pe/ps)。
2. 必须新建 → 建**时序表**(键含 trade_date/year/ann_date),**禁建快照表**。
3. 写入脚本归 `data/`(段2),读取归 `features/loaders.py`(段3,PIT 切)。
4. 更新本文件 + `FEATURE_CATALOG.md`。
