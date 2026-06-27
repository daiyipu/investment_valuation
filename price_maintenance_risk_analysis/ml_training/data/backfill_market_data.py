"""为指定样本的股票回填行情表(qfq/daily_basic/monthly)到接近全覆盖。
解决: Stage2 prefetch 限流失败留缺口 → 派生阶段逐行 pro_bar 回退 → 卡死。
用法: python backfill_market_data.py --samples <pilot_samples_*.parquet> [--passes 5]
读样本唯一股 → 多趟 prefetch(本地优先+缺失tushare落盘)→ 每趟报覆盖率, 直到 ≥99% 或趟数用尽。"""
import argparse, os, sys, time, warnings
warnings.filterwarnings('ignore')
PKG = '/Users/davy/github/investment_valuation/price_maintenance_risk_analysis'
os.chdir(PKG); sys.path.insert(0, PKG); sys.path.insert(0, os.path.join(PKG, 'ml_training'))
import pandas as pd, pymysql
from features.derive_features import prefetch_ohlcv, prefetch_daily_basic, prefetch_monthly
from utils.db_manager import ValuationDB

ap = argparse.ArgumentParser()
ap.add_argument('--samples', required=True)
ap.add_argument('--passes', type=int, default=5)
args = ap.parse_args()

codes = pd.read_parquet(args.samples)['股票代码'].astype(str).unique().tolist()
print(f'回填 {len(codes)} 股的行情表(qfq/daily_basic/monthly), 最多 {args.passes} 趟')

cfg = ValuationDB.MYSQL_CONFIG
def coverage():
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'], password=cfg['password'],
                           database=cfg['database'], charset=cfg['charset'])
    ph = ','.join(['%s'] * len(codes))
    out = {}
    for t in ['stock_qfq_daily', 'stock_daily_basic', 'stock_monthly_hfq']:
        out[t] = int(pd.read_sql(f'SELECT COUNT(DISTINCT stock_code) n FROM {t} WHERE stock_code IN ({ph})',
                                 conn, params=codes).iloc[0, 0])
    conn.close()
    return out

cov = coverage()
print(f'起点: qfq {cov["stock_qfq_daily"]}/{len(codes)} | daily_basic {cov["stock_daily_basic"]}/{len(codes)} | monthly {cov["stock_monthly_hfq"]}/{len(codes)}')

for i in range(args.passes):
    done = all(cov[t] >= 0.99 * len(codes) for t in cov)
    if done:
        print('全部 ≥99%, 完成'); break
    print(f'\n=== 趟 {i+1}/{args.passes} ===')
    t0 = time.time()
    for fn, name in [(prefetch_ohlcv, 'qfq'), (prefetch_daily_basic, 'daily_basic'), (prefetch_monthly, 'monthly')]:
        try:
            fn(codes)
        except Exception as e:
            print(f'  {name} 异常: {e}')
    cov = coverage()
    print(f'  覆盖({time.time()-t0:.0f}s): qfq {cov["stock_qfq_daily"]}({cov["stock_qfq_daily"]/len(codes)*100:.0f}%) | '
          f'daily_basic {cov["stock_daily_basic"]}({cov["stock_daily_basic"]/len(codes)*100:.0f}%) | '
          f'monthly {cov["stock_monthly_hfq"]}({cov["stock_monthly_hfq"]/len(codes)*100:.0f}%)')

print(f'\n✅ 终态: qfq {cov["stock_qfq_daily"]} | daily_basic {cov["stock_daily_basic"]} | monthly {cov["stock_monthly_hfq"]} / {len(codes)}')
