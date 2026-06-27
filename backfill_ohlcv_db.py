"""内存友好的 qfq+daily_basic 补全: 分批 fetch+persist+清缓存(~0.1GB/批), 8GB Mac 安全。
补全后 panel 构建全程读本地, 零 tushare + 低内存。"""
import sys, os, warnings
warnings.filterwarnings('ignore')
PKG = '/Users/davy/github/investment_valuation/price_maintenance_risk_analysis'
sys.path.insert(0, PKG); sys.path.insert(0, os.path.join(PKG, 'ml_training'))
os.chdir(PKG)
import pandas as pd, time
from derive_features import prefetch_ohlcv, prefetch_daily_basic, _OHLCV_CACHE, _DAILY_BASIC_CACHE
import derive_features as D

codes = pd.read_parquet('ml_training/data/backtest_samples.parquet')['股票代码'].astype(str).unique().tolist()
print(f'补全 {len(codes)} 股的 qfq + daily_basic(分批 300, 清缓存控内存)...')

BATCH = 300
t0 = time.time()
for i in range(0, len(codes), BATCH):
    b = codes[i:i + BATCH]
    prefetch_ohlcv(b)           # local 命中则秒过; 缺失则 fetch+落盘
    D._OHLCV_CACHE.clear()      # 释放(~0.13GB/批)
    prefetch_daily_basic(b)
    D._DAILY_BASIC_CACHE.clear()
    if (i // BATCH + 1) % 5 == 0:
        print(f'  {i+len(b)}/{len(codes)} ({time.time()-t0:.0f}s)')

# 终态
import pymysql
from utils.db_manager import ValuationDB
c = ValuationDB.MYSQL_CONFIG
conn = pymysql.connect(host=c['host'], port=c['port'], user=c['user'], password=c['password'], database=c['database'], charset=c['charset'])
for t in ['stock_qfq_daily', 'stock_daily_basic']:
    n = pd.read_sql(f'SELECT COUNT(DISTINCT stock_code) n FROM {t}', conn).iloc[0, 0]
    print(f'  {t}: {n}/{len(codes)} 股')
conn.close()
print(f'✅ 补全完成 ({time.time()-t0:.0f}s)')
