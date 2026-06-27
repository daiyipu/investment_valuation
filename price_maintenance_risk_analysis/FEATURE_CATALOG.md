# FEATURE_CATALOG — 全特征主目录（一处维护全特征的"说明书"）

> **权威单一源**:每个特征一行——类别 / 来源表 / 数据源 / PIT 机制 / 业务解释 / 产出脚本 / 消费模型。
> 新增字段必登记;`assert_panel_superset` 把关(ml_model_meta 并集 ⊆ panel 列)。
> **新增字段标准流程(铁律)**:① 查本目录定类别+归属脚本 → ② 在**现有脚本**里加(不开新表/新脚本;需新数据源先在 `data/` 落**时序表**)→ ③ 更新该脚本头部 docstring + 本目录登记一行。
> 产出脚本现状:`loaders`=export_features.py / `derive`=derive_features.py / `operators`=factor_engine.py(Phase 2 后迁 features/)。

---

## ① 行情类（~33 个）— loaders/derive · stock_qfq_daily · PIT searchsorted≤date
**产出**:`derive.derive_market_stats_from_ohlcv`(基列)+ `derive.derive_market_momentum`(派生比)。

| 特征 | 定义 | 来源表 | PIT | 产出脚本 | 消费模型 |
|---|---|---|---|---|---|
| MA20/30/60/120/250 | close.rolling(w).mean | stock_qfq_daily | searchsorted≤date | derive_market_stats | 1w(历史)/参考 |
| 波动率_20/30/60/120/250d | log收益 rolling(w).std×√250 | stock_qfq_daily | 同 | derive_market_stats | 多 |
| 年化收益_20/30/60/120/250d | (close/close.shift(w))^(250/w)-1 | stock_qfq_daily | 同 | derive_market_stats | 多 |
| 区间收益_*d | close/close.shift(w)-1 | stock_qfq_daily | 同 | derive_market_stats | 多 |
| 胜率_*d | rolling(w).mean(ret>0) | stock_qfq_daily | 同 | derive_market_stats | — |
| 当前价 / 漂移率 / 波动率(全期) | close / expanding.mean×250 / expanding.std×√250 | stock_qfq_daily | 同 | derive_market_stats | — |
| 中位价 / 价格标准差 / 数据天数 | expanding.median / expanding.std / count | stock_qfq_daily | 同 | derive_market_stats | — |
| vol_ratio_20_60 / vol_ratio_60_250 | 波动率短/长比 | (派生) | 同 | derive_market_momentum | — |
| return_acceleration | 年化收益_60d − 年化收益_120d | (派生) | 同 | derive_market_momentum | 1w(历史) |
| price_vs_MA20/60/120/250 | 当前价/MA | (派生) | 同 | derive_market_momentum | 1w(历史) |

## ② 估值类（个股/行业/同行 PE/PB/PS + _vs_*）— loaders · stock_daily_basic + industry_daily · PIT
**产出**:`loaders.load_pb_vs_industry`(PB PIT)+ `derive.derive_valuation_relative`(vs 派生)。⚠️ PE/PS 个股+行业 PIT loader 待 Phase 1 新增(原 load_db_features 读快照非PIT,退役)。

| 特征 | 定义 | 来源表 | PIT | 产出脚本 | 状态 |
|---|---|---|---|---|---|
| 个股PB | daily_basic.pb ≤date | stock_daily_basic | searchsorted | load_pb_vs_industry | ✅ |
| 个股PE / 个股PS | daily_basic.pe/ps ≤date | stock_daily_basic(待补列) | searchsorted | **待新增 loader** | ⏳ Phase1 |
| 行业PE / 行业PB | industry_daily pe/pb ≤date | industry_daily | searchsorted | **待新增 loader** | ⏳ Phase1 |
| 同行PE/PB/PS_均值/中位 | 同行业成分 ≤date 聚合(或行业指数代理) | industry_daily / daily_basic | 截面聚合 | **待新增 loader** | ⏳ Phase1 |
| PB_vs_同行中位(PIT行业口径) | 个股PB/行业PB | (派生) | PIT | derive_pb_vs_industry_pit | ✅(7m 模型用) |
| PE_vs_行业 / PB_vs_行业 | 个股/行业 | (派生) | — | derive_valuation_relative | ⏳(等个股PE loader) |
| PE_vs_同行均值/中位 / PB_vs_同行均值 / PS_vs_同行均值 | 个股/同行 | (派生) | — | derive_valuation_relative | ⏳ |

## ③ 财务类（FCF + 评分 + 27 比率）— loaders · historical_fcf + company_annual_scores + financial_indicators · PIT year≤pit_year
**产出**:`loaders.load_fcf_bulk`/`load_fcf_accel`/`load_total_score_delta2y`/`load_financial_ratios` + `derive.derive_fcf_*`/`derive_financial_score_deltas` + build_backtest_panel `_bulk_score_base`。

| 特征族 | 定义 | 来源表 | PIT | 产出脚本 | 消费模型 |
|---|---|---|---|---|---|
| FCF T/T1/T2/T3/T4(营收/NOPAT/FCF/营业利润/净利润_fcf/资本支出/折旧/营运资金变动) | year≤pit_year 取近5年 | historical_fcf | year≤pit_year | load_fcf_bulk | 多 |
| {metric}_YoY / _CAGR2 / _加速 | YoY / 2年CAGR / YoY加速度 | (派生) | 同 | derive_fcf_growth_rates | 7m(FCF_加速) |
| FCF_margin / FCF_conversion / capex_to_dep / capex_intensity / NOPAT_margin | FCF比率 | (派生) | 同 | derive_fcf_cross_metrics | — |
| 总分/盈利能力/成长能力 × T/T-1/T-2/T-3/T-4 | report_year≤pit_year 取近5年 | company_annual_scores | report_year≤pit_year | _bulk_score_base | 多 |
| 总分_delta_1y/2y/4y | T−T-k | (派生) | 同 | derive_financial_score_deltas | 7m(总分_delta_2y) |
| 27 财务比率(流动/速动/周转/ROA/ROE/净利率/资产负债/…) | ann_date≤date | financial_indicators | ann_date≤date | load_financial_ratios | 7m(应收周转/产权/利息资本化) |
| 净资产增长 | (派生) | financial_indicators | 同 | derive | 7m |
| 总分_斜率 / 成长能力_斜率 | polyfit T..T-4 序列 | company_annual_scores | year≤pit_year | **待新增(Phase1 Tier2)** | ⏳(1w等历史用) |

## ④ 因子类（因子引擎 ~56 个）— operators/derive · stock_qfq_daily + daily_basic + 行业/大盘 · PIT(derive_alpha_beta 切≤D 序列喂 compute_factors_series)
**产出**:`operators.factor_engine.compute_factors_series` + `derive.derive_alpha_beta_factors`。

| 族 | 代表特征 | 来源 | 模型用 |
|---|---|---|---|
| kline | k_KMID/KLEN/KMID2/KUP/KLOW/k_BODY_RATIO | close/open/high/low | 1w(k_BODY_RATIO/KMID2 历史) |
| tech | RSI_6/12/24 / KDJ_K/D/J / MACD_DIF/DEA/HIST / BOLL_%B/BW | close | — |
| volume | corr_ret_vol_n / vol_* / amount_* / vwap / obv / pvt / mfi / cmf | close/vol/amount | 7m(pvt_slope_20) |
| moment | ret_skew_60 / ret_kurt_60 / ROC_n | close | 7m(ret_skew_60) / 1w(ret_kurt_60) |
| beta | beta_mkt_* / idiovol | 三序列对齐收益 | — |
| turnover | turnover_mean_20 / turnover_std/skew / turnover_now | daily_basic | 7m(turnover_mean_20) / 1w(turnover_now) |
| multiperiod | MACD_W_HIST / ROC_W_3(周线重采样≤D) | close 周线 | 7m(MACD_W_HIST/ROC_W_3) |
| smc | smc_premium/bos/liq_sweep/ote/displacement… | close/high/low | — |

## ⑤ 行业估值增长 / 大盘 / 月线类 — derive · industry_daily + index_daily · PIT
**产出**:`derive.derive_industry_valuation_growth`/`derive_market_index_features`/`derive_monthly_trend`/`derive_market_trend`。

| 特征 | 定义 | 来源 | 消费模型 |
|---|---|---|---|
| 行业PE/PB_60/120/250d增长 | 行业指数 pe/pb 增长率 | industry_daily | 7m(行业PE_250d增长) |
| 市场_above_MA250 / 距离MA250 / 波动率_120d / 年化收益_120d / 胜率_120d / 波动率比值 | 大盘沪深300 统计 | index_daily | — |
| 月线MA10_slope3% / 月线趋势向上 | 月线 close 趋势 | stock_monthly_hfq | — |
| 大盘MA10_slope3% | 沪深300 趋势 | index_daily | — |

## ⑥ 策略类（三浪/抵抗）— derive.derive_strategy_signals · stock_qfq_daily + 行业/大盘 · PIT
| 三浪_score/gain/retr | 抵抗_score/corr_div_stock | 波二+抵抗 PIT 信号 | derive_strategy_signals |

## ⑦ SUE 类（业绩超预期）— loaders · sue_timelines.parquet(fetch_factors._sue_for_sample) · PIT ann_date≤date
| 特征 | 定义 | 状态 |
|---|---|---|
| sue_beat | (实际/快报净利 − 预告中点)/\|预告中点\| | ✅ load_sue_beat(7m) |
| sue_zscore / sue_pos_streak / sue_recency_d / sue_up_trend / sue_yoy_acc / sue_yoy_mean3 | SUE 全系列(原 placement_evaluation 回填) | ⏳ Phase1 Tier2(原特征名,1w/3m 模型用) |

## ⑧ 筹码类（chip）— loaders · cyq_chips · PIT ⚠️仅2018+覆盖
| chip_winner_rate / chip_avg_cost_dev / chip_concentration / chip_peak_dev / chip_cost_spread | 筹码分布 PIT | ⏳ Phase1 Tier2(原特征名,早年NaN) |

## ⑨ 资金流类（mf）— loaders · moneyflow · PIT
| mf_main_net_ratio_5d/20d / mf_net_mf_ratio_20d / mf_main_mom / mf_sm_net_ratio_20d | 资金流 PIT | ⏳ Phase1 Tier2 |

## ⑩ 持股类（nb_hold）— loaders · pro.hk_hold · PIT trade_date≤date
| nb_hold_ratio | 沪深港通持股比 ≤date | ✅ load_nb_hold(7m) |
| nb_hold_chg_20d / nb_hold_chg_60d | 持股比变化 | ⏳ Phase1 Tier2(1w历史) |

## ⑪ 定增结构类（placement, 回测 skip）— derive.derive_placement_structure
| 折价率 / 定增稀释率 / 募集市值比 / 定增大股东参与 / 定增锁定期天数 | 定增专属,全A 回测 skip_placement | — |

---

## 恢复优先级（Phase 1）
- ✅ 已恢复(本次): 行情类 ~33(MA/波动率/年化收益/…) via derive_market_stats_from_ohlcv;T-3 评分。
- ⏳ Tier1 待办: 估值 PE/PS 个股+行业+同行 PIT loader(喂 derive_valuation_relative 产 _vs_*)。
- ⏳ Tier2 待办(原特征名,现有 1m/3m/1w/2w 模型直接跑): sue_zscore 等6个 / chip 5个 / mf 5个 / nb_hold_chg 2个 / 总分_斜率·成长能力_斜率 2个。
- 铁律:恢复用**原特征名**,现有模型不重训即可打分。

## 消费模型速查
- **7m 生产模型(15feat)**:FCF_加速 / MACD_W_HIST / 净资产增长 / pvt_slope_20 / turnover_mean_20 / nb_hold_ratio / PB_vs_同行中位 / 应收账款周转率 / 产权比率 / 利息资本化率 / ROC_W_3 / ret_skew_60 / sue_beat / 总分_delta_2y / 行业PE_250d增长。
- **短周期(1m/3m/1w/2w)历史模型**:含 sue_zscore/行业年化收益_*/nb_hold_chg_*/成长能力_斜率/同行PE_均值/MA*/k_BODY_RATIO 等(Phase1 恢复后可直接跑)。
