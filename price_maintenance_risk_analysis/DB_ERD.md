# 数据库实体关系图（ERD）— investment_valuation 库

> **治理目的:防止乱设库表 / 重复表 / 主键错位。** 新增表前先查本图:能否落已有表(加列)、能否复用已有自然键、属时序还是快照。
> 互链:[DB_SCHEMA.md](DB_SCHEMA.md)(每表列清单 + 时序/快照标记)、[FEATURE_CATALOG.md](FEATURE_CATALOG.md)(特征→表映射)。
> 共 **27 张表**。关系以**自然键**(stock_code / trade_date / issue_date / index_code / version / report_year)为准,代理 `id` 仅作行号。

---

## 一、领域划分(5 域)

| 域 | 表 | 角色 |
|---|---|---|
| **A. 个股行情/估值(时序,PIT 可切)** | `stock_qfq_daily` `stock_daily_basic` `stock_monthly_hfq` | 特征层唯一行情源 |
| **B. 行业/指数(时序)** | `industry_daily` `index_factor_pro` `industry_data` `industry_financials` `em_industry_boards` `em_industry_stocks` | 行业映射 + 行业/指数 PE/PB/技术面 |
| **C. 个股财务(时序,年报键)** | `historical_fcf` `company_annual_scores` `financial_indicators` | FCF / 评分 / 27 比率 |
| **D. 定增脊柱+标签+快照** | `placement_evaluation` `issue_date_locked` `placement_params` `market_data` `relative_valuation` `peer_companies` `market_indices` `screening_results` `stocks` | 定增训练脊柱 + 🔴退役快照 |
| **E. ML 元数据/产物** | `ml_model_meta` `ml_dataset_snapshot` `ml_validation` `ml_validation_byyear` `ml_validation_groups` `ml_validation_sections` | 模型/特征数据集/验证结果(常态产出入库于此)。`ml_panel`(回测panel BLOB)规划中,待 MySQL 配置 |

---

## 二、ER 图（Mermaid）

```mermaid
erDiagram
    stocks ||--o{ stock_qfq_daily : "stock_code"
    stocks ||--o{ stock_daily_basic : "stock_code"
    stocks ||--o{ stock_monthly_hfq : "stock_code"
    stocks ||--o{ placement_evaluation : "stock_code"
    stocks ||--o{ historical_fcf : "stock_code"
    stocks ||--o{ company_annual_scores : "stock_code"
    stocks ||--o{ financial_indicators : "stock_code"
    stocks ||--o{ industry_data : "stock_code→index_code"
    stocks ||--o{ issue_date_locked : "stock_code"
    stocks ||--o{ placement_params : "stock_code"
    stocks ||--o| market_data : "stock_code 🔴快照"
    stocks ||--o| relative_valuation : "stock_code 🔴快照"
    stocks ||--o{ peer_companies : "stock_code 🔴快照"
    stocks ||--o{ screening_results : "stock_code 🔴快照"

    placement_evaluation ||--|| issue_date_locked : "(stock_code,issue_date)"
    placement_evaluation }o--|| placement_params : "stock_code"

    industry_data }o--|| industry_daily : "index_code,trade_date"
    industry_data }o--|| index_factor_pro : "index_code,trade_date"
    industry_daily ||--o{ index_factor_pro : "同 index_code+trade_date"

    historical_fcf }o..|| company_annual_scores : "stock_code+year≈report_year(弱)"
    financial_indicators }o..|| company_annual_scores : "stock_code+report_year(弱)"

    ml_model_meta ||--o{ ml_validation : "version"
    ml_model_meta ||--o{ ml_validation_byyear : "version"
    ml_model_meta ||--o{ ml_validation_groups : "version"
    ml_model_meta ||--o{ ml_validation_sections : "version"
    ml_model_meta ||--|| ml_dataset_snapshot : "version"

    ml_panel {
        string tag PK
        string kind
        int n_rows n_cols size_mb
        string sha256 UK
        longblob parquet
        note string "⏳规划中, 待MySQL max_allowed_packet配好"
    }

    stocks {
        string stock_code PK
        string stock_name
        string industry_code
    }
    stock_qfq_daily {
        string stock_code PK
        string trade_date PK
        double open high low close vol amount
    }
    stock_daily_basic {
        string stock_code PK
        string trade_date PK
        double turnover volume_ratio pb pe ps total_mv
    }
    placement_evaluation {
        bigint id PK
        string stock_code UK
        string issue_date UK
        double return_1w_2w_4w_1m_3m_7m
        double chip_mf_smc_nb_sue_pp
    }
    issue_date_locked {
        string stock_code PK
        string issue_date PK
        double issue_date_price ma20
    }
    industry_daily {
        bigint id PK
        string index_code UK
        string trade_date UK
        string data_source
        double pe pb close
    }
    index_factor_pro {
        string index_code PK
        string trade_date PK
        double MACD_KDJ_RSI_BOLL_87因子
    }
    historical_fcf {
        bigint id PK
        string stock_code UK
        int year UK
        double revenue nopat fcf capex
    }
    company_annual_scores {
        bigint id PK
        string stock_code UK
        int report_year UK
        double total profit growth
    }
    ml_model_meta {
        string version PK
        string kind horizon label_config
        json features
        blob lr_bundle
    }
    ml_validation_sections {
        bigint id PK
        string version UK
        string date
        double ic long short ls
    }
    ml_dataset_snapshot {
        string version PK
        string sha256
        longblob parquet
    }
```

> 🔴 = 快照表(非 PIT,特征层退役中,仅报告用)。`UK` = 应有但当前未强制的自然唯一键(见下节治理)。

---

## 三、关系矩阵（自然键 join）

| 父键 | 共享该键的表 | 关系类型 |
|---|---|---|
| `stock_code` | stocks → stock_qfq_daily / stock_daily_basic / stock_monthly_hfq / placement_evaluation / historical_fcf / company_annual_scores / financial_indicators / industry_data / issue_date_locked / placement_params / market_data / relative_valuation / peer_companies / screening_results | 1对多 |
| `(stock_code, trade_date)` | stock_qfq_daily · stock_daily_basic · stock_monthly_hfq | 三表同键(个股日线三视图) |
| `(stock_code, issue_date)` | placement_evaluation · issue_date_locked | 1对1(报价日精确取价) |
| `(stock_code, year/report_year)` | historical_fcf(year) · company_annual_scores(report_year) · financial_indicators(report_year) | 年度财务三视图(report_year≠calendar year,按 ann_date PIT) |
| `(index_code, trade_date)` | industry_daily · index_factor_pro | 行业指数两视图(基本面 + 技术面) |
| `index_code` | industry_data(stock→index 映射) → industry_daily / index_factor_pro | 映射桥 |
| `version` | ml_model_meta → ml_dataset_snapshot / ml_validation / ml_validation_byyear / ml_validation_groups / ml_validation_sections | 模型 1 对 多产物 |

---

## 四、治理标记（防乱建库表）

### ⚠️ 1. 代理 `id` 主键但自然键未强制(易插重复行)
下列表 PK 是 `id`,但业务唯一性靠**自然键**,当前**无 UNIQUE 约束** → 同一 (stock_code, date/year) 可被重复插入(已发生过:save-NaN-silent-drop-bug 类)。**建议加 UNIQUE 索引**:

| 表 | 自然唯一键 | 现状 |
|---|---|---|
| placement_evaluation | (stock_code, issue_date) | id PK, 无 UNIQUE ⚠️ |
| historical_fcf | (stock_code, year) | id PK, 无 UNIQUE ⚠️ |
| company_annual_scores | (stock_code, report_year) | id PK, 无 UNIQUE ⚠️ |
| financial_indicators | (stock_code, report_year) | id PK, 无 UNIQUE ⚠️ |
| industry_daily | (index_code, trade_date, data_source) | id PK, 无 UNIQUE ⚠️ |
| peer_companies | (stock_code, peer_code) | id PK, 无 UNIQUE |
| screening_results | (stock_code, step) | id PK, 无 UNIQUE |
| ml_validation* (4表) | (version, …) | id PK, 应挂 version FK |

### 🔴 2. 快照表(非 PIT,退役中 — 勿再扩,勿喂特征)
`market_data` `relative_valuation` `peer_companies` `market_indices` `screening_results` — 每股/每指数一行无 trade_date 维度,读即可能引入报价日后信息(泄漏)。特征层已停读,改走时序表 PIT 重算(见 DB_SCHEMA.md 退役去向)。

### ♻️ 3. 功能重叠(避免再建类似表)
- `market_data`(快照行情统计) ⊊ `stock_qfq_daily` PIT 重算(`derive_market_stats_from_ohlcv`)。
- `relative_valuation` + `peer_companies`(快照估值) ⊊ `stock_daily_basic` + `industry_daily` PIT(`derive_peer_valuation` / `load_pb_vs_industry`)。
- `market_indices`(locked_date 快照) ⊊ 运行时 `index_daily` PIT(`derive_market_index_features`)。
- `industry_financials`(空表 0 行) vs `industry_daily` pe/pb — 前者空,勿用。

### 📥 4. 常态产出(验证 csv / 特征 panel)应入库,不裸露文件
- **验证结果(逐截面/年/组 IC-L-S)** → 入 `ml_validation*` 4 表,唯一入口 `save_validation_db.save_validation_run(df,...)`(validate_5h / backtest_long_short 直调,**禁再产裸 csv**)。
- **特征 panel(large parquet)** → ⏳ **落库延后**。正道 = `save_validation_db.register_panel(path,tag,...)` 以 BLOB 入专用表 `ml_panel`(tag/sha256/n_rows/n_cols/parquet LONGBLOB, 入库后删裸文件)。⚠️ 前置: 需先 `SET GLOBAL max_allowed_packet=536870912` + 重启 MySQL(当前 64MB < 大 panel 241MB, 超包报错)。**配置前 panel 暂留 `ml_training/data`, register_panel 已就位待用, 不强制现在落库。**
- **模型权重** → 已在 `ml_model_meta.lr_bundle`(BLOB),磁盘 output/v_*/ 仅 fallback。
- **特征数据集快照(base/derived features parquet)** → `ml_dataset_snapshot`(kind enum base/derived, sha256, 小文件 LONGBLOB)。注意:勿与回测 panel 混(后者走 `ml_panel`)。

---

## 五、新增表/列的决策树

```
要加字段/表?
├─ 能否落已有表(加列)?  → 是: 加列(时序表优先), 更新 DB_SCHEMA.md + FEATURE_CATALOG.md
│                        否 ↓
├─ 是时序(有 trade_date/year/ann_date)?  → 建**时序表**, 自然复合键作 PK(勿用 id)
│   否(快照/每股一行) ↓
├─ 必须吗? 定增专属/报告辅助? → 是: 标 🔴快照, 仅报告用, 不喂特征
│                          否 ↓
└─ ⛔ 停。多半是重复表(查 §3 重叠)。改走已有时序表 PIT 重算。
```

---

## 六、取数层脚本归属（防脚本散乱 — Phase2 落地准绳）

**铁律:一张表 = 一个数据域模块里的一个写者函数。新数据源/新表 → 定域 → 加进 `data/<域>.py`,绝不新开脚本。**(三层一入口·取数层; 用户 2026-06-27 确认)

当前 ~12 个散落写库脚本 → Phase2 归并为 **6 个域模块 + 1 编排**:

| 数据域 | 表 | 当前散落脚本 | Phase2 目标模块 |
|---|---|---|---|
| 个股行情 | stock_qfq_daily / stock_daily_basic / stock_monthly_hfq | update_market_data, backfill_market_data, ingest_raw, derive_features(`_qfq_save/_db_basic_save/_monthly_save`) | `data/market.py` |
| 行业/指数 | industry_daily / index_factor_pro / industry_data | refresh_industry_daily, refresh_industry_mapping | `data/industry.py` |
| 财务 | historical_fcf / company_annual_scores / financial_indicators | backfill_financial_indicators, batch_financial_score, bulk_backfill_early_scores | `data/financial.py` |
| 定增因子+标签 | placement_evaluation | fetch_factors, compute_labels, backfill_evaluations | `data/factors.py` + `data/labels.py` |
| ML 元数据/产物 | ml_model_meta / ml_dataset_snapshot / ml_validation* | db_model_store, db_dataset_store, save_validation_db | `data/ml_meta.py` |
| 编排 | — | update_all_data | `data/orchestrate.py` |

**新增数据源的决策树:**
```
要灌新数据源/新表?
├─ 属哪个数据域?(行情/行业/财务/因子标签/ML元) → 定域
├─ 该域 data/<域>.py 里加 ingest/backfill 函数(落已有表加列, 或建时序表, 见 §五)
├─ update_all_data 编排里挂一步
└─ ⛔ 禁新开 scripts/xxx_backfill.py 散脚本(违反三层一入口)
```
Phase2 搬 `data/` 6 模块: `git mv` + 全项目 import 改 + 每模块搬完跑 smoke 再下一个(可中途停)。搬完后每表写者唯一、新需求有明确落点。

