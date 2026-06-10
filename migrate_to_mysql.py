-- MySQL建表脚本 - 定增估值数据库
-- 数据库: investment_valuation

CREATE DATABASE IF NOT EXISTS investment_valuation
    DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE investment_valuation;

-- 1. 股票主表
CREATE TABLE IF NOT EXISTS stocks (
    stock_code   VARCHAR(20) PRIMARY KEY,
    stock_name   VARCHAR(100),
    sw_l1_code   VARCHAR(20), sw_l1_name VARCHAR(50),
    sw_l2_code   VARCHAR(20), sw_l2_name VARCHAR(50),
    sw_l3_code   VARCHAR(20), sw_l3_name VARCHAR(50),
    created_at   DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at   DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 2. 定增参数
CREATE TABLE IF NOT EXISTS placement_params (
    stock_code       VARCHAR(20) PRIMARY KEY,
    financing_amount BIGINT DEFAULT 100000000,
    lockup_period    INT DEFAULT 6,
    pricing_method   VARCHAR(50) DEFAULT 'ma20_discount_90',
    premium_rate     DOUBLE DEFAULT -0.10,
    risk_free_rate   DOUBLE DEFAULT 0.03,
    net_assets       DOUBLE DEFAULT 0,
    total_debt       DOUBLE DEFAULT 0,
    net_income       DOUBLE DEFAULT 0,
    revenue_growth   DOUBLE DEFAULT 0.15,
    operating_margin DOUBLE DEFAULT 0.15,
    beta             DOUBLE DEFAULT 1.0,
    issue_date       VARCHAR(10),
    issue_price      DOUBLE,
    issue_shares     BIGINT,
    current_price    DOUBLE,
    created_at       DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at       DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 3. 市场数据
CREATE TABLE IF NOT EXISTS market_data (
    stock_code       VARCHAR(20) PRIMARY KEY,
    analysis_date    VARCHAR(10),
    latest_trading_date VARCHAR(10),
    issue_date       VARCHAR(10), invitation_date VARCHAR(10),
    current_price    DOUBLE, avg_price_all DOUBLE, median_price DOUBLE, price_std DOUBLE,
    volatility_20d   DOUBLE, volatility_60d DOUBLE, volatility_120d DOUBLE, volatility_250d DOUBLE,
    annual_return_20d DOUBLE, annual_return_60d DOUBLE, annual_return_120d DOUBLE, annual_return_250d DOUBLE,
    period_return_20d DOUBLE, period_return_60d DOUBLE, period_return_120d DOUBLE, period_return_250d DOUBLE,
    ma_20 DOUBLE, ma_30 DOUBLE, ma_60 DOUBLE, ma_120 DOUBLE, ma_250 DOUBLE,
    win_rate_20d DOUBLE, win_rate_60d DOUBLE, win_rate_120d DOUBLE, win_rate_250d DOUBLE,
    total_days INT, drift DOUBLE, volatility DOUBLE,
    price_series     LONGTEXT,
    market_turnover  LONGTEXT,
    schema_version   VARCHAR(30),
    data_source      VARCHAR(30) DEFAULT 'tushare_realtime',
    generated_at     DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 4. 行业指数数据
CREATE TABLE IF NOT EXISTS industry_data (
    stock_code    VARCHAR(20) PRIMARY KEY,
    index_code    VARCHAR(20),
    industry_name VARCHAR(100),
    sw_l1_code VARCHAR(20), sw_l1_name VARCHAR(50),
    sw_l2_code VARCHAR(20), sw_l2_name VARCHAR(50),
    sw_l3_code VARCHAR(20), sw_l3_name VARCHAR(50),
    analysis_date  VARCHAR(10), current_level DOUBLE,
    volatility_20d DOUBLE, volatility_60d DOUBLE, volatility_120d DOUBLE, volatility_250d DOUBLE,
    annual_return_20d DOUBLE, annual_return_60d DOUBLE, annual_return_120d DOUBLE, annual_return_250d DOUBLE,
    period_return_20d DOUBLE, period_return_60d DOUBLE, period_return_120d DOUBLE, period_return_250d DOUBLE,
    ma_20 DOUBLE, ma_60 DOUBLE, ma_120 DOUBLE, ma_250 DOUBLE,
    win_rate_20d DOUBLE, win_rate_60d DOUBLE, win_rate_120d DOUBLE, win_rate_250d DOUBLE,
    total_days INT, drift DOUBLE, volatility DOUBLE,
    data_source VARCHAR(30) DEFAULT 'tushare_sw_index',
    generated_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 5. 行业日线
CREATE TABLE IF NOT EXISTS industry_daily (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    index_code      VARCHAR(20) NOT NULL,
    trade_date      VARCHAR(10) NOT NULL,
    open DOUBLE, high DOUBLE, low DOUBLE, close DOUBLE,
    volume DOUBLE, amount DOUBLE, pct_chg DOUBLE,
    pe DOUBLE, pb DOUBLE, ps_ttm DOUBLE,
    data_source     VARCHAR(20) DEFAULT 'tushare_sw',
    UNIQUE KEY uk_ind_daily (index_code, trade_date, data_source)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 6. 相对估值
CREATE TABLE IF NOT EXISTS relative_valuation (
    stock_code        VARCHAR(20) PRIMARY KEY,
    cache_date        VARCHAR(10) NOT NULL,
    trade_date        VARCHAR(10),
    current_pe DOUBLE, current_pb DOUBLE, current_ps DOUBLE,
    sw_index_pe DOUBLE, sw_index_pb DOUBLE, sw_index_ps DOUBLE,
    target_index_code VARCHAR(20), target_industry_l3 VARCHAR(100),
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 7. 同行公司
CREATE TABLE IF NOT EXISTS peer_companies (
    id          INT AUTO_INCREMENT PRIMARY KEY,
    stock_code  VARCHAR(20) NOT NULL,
    peer_name   VARCHAR(100), peer_code VARCHAR(20),
    pe DOUBLE, ps DOUBLE, pb DOUBLE, market_cap DOUBLE,
    INDEX idx_peer_stock (stock_code)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 8. 锁定日
CREATE TABLE IF NOT EXISTS issue_date_locked (
    stock_code      VARCHAR(20) NOT NULL,
    issue_date      VARCHAR(10) NOT NULL,
    issue_date_price DOUBLE,
    ma_20           DOUBLE,
    current_price   DOUBLE,
    analysis_date   VARCHAR(10),
    locked_timestamp VARCHAR(30),
    PRIMARY KEY (stock_code, issue_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 9. 历史FCF
CREATE TABLE IF NOT EXISTS historical_fcf (
    id          INT AUTO_INCREMENT PRIMARY KEY,
    stock_code  VARCHAR(20) NOT NULL,
    year        INT NOT NULL,
    revenue DOUBLE, operate_profit DOUBLE, net_income DOUBLE,
    nopat DOUBLE, depreciation DOUBLE, capex DOUBLE,
    wc_change DOUBLE, fcf DOUBLE,
    UNIQUE KEY uk_fcf (stock_code, year)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 10. 市场指数
CREATE TABLE IF NOT EXISTS market_indices (
    index_code  VARCHAR(20) NOT NULL,
    index_name  VARCHAR(50) NOT NULL,
    locked_date VARCHAR(10) NOT NULL DEFAULT '',
    current_level DOUBLE,
    volatility_20d DOUBLE, volatility_60d DOUBLE, volatility_120d DOUBLE, volatility_250d DOUBLE,
    return_20d DOUBLE, return_60d DOUBLE, return_120d DOUBLE, return_250d DOUBLE,
    period_log_return_20d DOUBLE, period_log_return_60d DOUBLE,
    period_log_return_120d DOUBLE, period_log_return_250d DOUBLE,
    ma_20 DOUBLE, ma_60 DOUBLE, ma_120 DOUBLE, ma_250 DOUBLE,
    win_rate_20d DOUBLE, win_rate_60d DOUBLE, win_rate_120d DOUBLE, win_rate_250d DOUBLE,
    data_date VARCHAR(10),
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (index_code, locked_date)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 11. 筛选结果
CREATE TABLE IF NOT EXISTS screening_results (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    batch_id        VARCHAR(30) NOT NULL,
    stock_code      VARCHAR(20),
    stock_name      VARCHAR(100),
    premium_min DOUBLE, premium_max DOUBLE,
    valid_thresholds INT,
    step1_pass INT, step1_detail TEXT,
    step2_pass INT, step2_detail TEXT,
    step3_pass INT, step3_detail TEXT,
    decision TEXT,
    summary LONGTEXT,
    error TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_batch (batch_id),
    INDEX idx_stock (stock_code)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 12. 东财板块
CREATE TABLE IF NOT EXISTS em_industry_boards (
    board_code     VARCHAR(20) PRIMARY KEY,
    board_name     VARCHAR(100) NOT NULL,
    total_count    INT DEFAULT 0,
    latest_price   DOUBLE, change_pct DOUBLE,
    total_mv       DOUBLE, turnover_rate DOUBLE,
    up_count INT, down_count INT,
    leading_stock  VARCHAR(100), leading_pct DOUBLE,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 13. 东财成份股
CREATE TABLE IF NOT EXISTS em_industry_stocks (
    id             INT AUTO_INCREMENT PRIMARY KEY,
    board_code     VARCHAR(20) NOT NULL,
    stock_code     VARCHAR(20) NOT NULL,
    stock_name     VARCHAR(100) NOT NULL,
    latest_price   DOUBLE, change_pct DOUBLE, change_amt DOUBLE,
    volume DOUBLE, amount DOUBLE, amplitude DOUBLE,
    high DOUBLE, low DOUBLE, open DOUBLE, prev_close DOUBLE,
    turnover_rate DOUBLE, pe_dynamic DOUBLE, pb DOUBLE,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY uk_em_stock (board_code, stock_code)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- 14. 行业财报(JSON缓存)
CREATE TABLE IF NOT EXISTS industry_financials (
    l3_code VARCHAR(20) PRIMARY KEY,
    l3_name VARCHAR(100),
    fetch_date VARCHAR(10),
    company_count INT,
    data_json LONGTEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
