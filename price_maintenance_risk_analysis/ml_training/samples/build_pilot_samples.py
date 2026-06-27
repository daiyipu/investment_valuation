"""建 pilot universe: N股 × 192月末 → pilot_samples_{mode}.parquet。
两种抽样(对比用, 测生存偏差+小盘倾斜):
  - stratified(默认): 申万L1行业比例分层(代表性, 但只抽有行业映射的股 → 边缘/退市股少)
  - random: 全universe(5302, 含退市/当前ST/无行业边缘股)均匀随机抽 → 成分更接近真实市场
8GB Mac 安全(panel ~96k行, OHLCV~0.2GB)。复用 backtest_samples 的月末报价日。"""
import sys, os, argparse, warnings
warnings.filterwarnings('ignore')
PKG = '/Users/davy/github/investment_valuation/price_maintenance_risk_analysis'
sys.path.insert(0, PKG); sys.path.insert(0, os.path.join(PKG, 'ml_training'))
os.chdir(PKG)
import pandas as pd, numpy as np, pymysql, tushare as ts
from utils.db_manager import ValuationDB
from tushare_token import resolve_tushare_token
os.environ.setdefault('TUSHARE_TOKEN', resolve_tushare_token())

ap = argparse.ArgumentParser()
ap.add_argument('--mode', choices=['stratified', 'random'], default='stratified')
ap.add_argument('--n', type=int, default=500)
ap.add_argument('--seed', type=int, default=42)
args = ap.parse_args()

# 全A样本 + 月末日期(零 placement 依赖)
samples = pd.read_parquet('ml_training/data/backtest_samples.parquet')
all_codes = samples['股票代码'].astype(str).unique()
dates = sorted(samples['报价日'].astype(str).unique())

if args.mode == 'stratified':
    # 行业映射(每股取最新) → L1比例分层
    cfg = ValuationDB.MYSQL_CONFIG
    conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                           password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
    ind = pd.read_sql('SELECT stock_code, sw_l1_name, analysis_date FROM industry_data '
                      'WHERE sw_l1_name IS NOT NULL ORDER BY analysis_date DESC', conn)
    conn.close()
    ind = ind.drop_duplicates('stock_code', keep='first')
    ind = ind[ind['stock_code'].astype(str).isin(all_codes)]
    print(f'有行业的universe股: {len(ind)} / {len(all_codes)}')
    rng = np.random.default_rng(args.seed)
    picked = []
    for name, g in ind.groupby('sw_l1_name'):
        k = max(1, round(len(g) * args.n / len(ind)))
        avail = g['stock_code'].astype(str).values
        picked += rng.choice(avail, size=min(k, len(avail)), replace=False).tolist()
    picked = list(dict.fromkeys(picked))[:args.n]
else:
    # 全universe均匀随机(含退市/ST/无行业), 不分层
    rng = np.random.default_rng(args.seed)
    picked = rng.choice(all_codes, size=min(args.n, len(all_codes)), replace=False).tolist()
    print(f'全universe随机抽: {len(picked)}/{len(all_codes)} (含退市/ST/无行业边缘股)')

picked = [str(c) for c in picked]
out = f'ml_training/data/pilot_samples_{args.mode}.parquet'

# 成分诊断(对比 stratified vs random 是否真有差异)
pro = ts.pro_api()
n_del = n_st = 0
try:
    dD = pro.stock_basic(list_status='D', fields='ts_code')
    n_del = len(set(dD['ts_code'].astype(str)) & set(picked))
except Exception:
    pass
try:
    dL = pro.stock_basic(list_status='L', fields='ts_code,name')
    st_codes = dL.loc[dL['name'].astype(str).str.contains('ST', na=False), 'ts_code'].astype(str)
    n_st = len(set(st_codes) & set(picked))
except Exception:
    pass
cfg = ValuationDB.MYSQL_CONFIG
conn = pymysql.connect(host=cfg['host'], port=cfg['port'], user=cfg['user'],
                       password=cfg['password'], database=cfg['database'], charset=cfg['charset'])
ph = ','.join(['%s'] * len(picked))
has_ind = pd.read_sql(f"SELECT DISTINCT stock_code FROM industry_data "
                      f"WHERE sw_l1_name IS NOT NULL AND stock_code IN ({ph})", conn, params=picked)['stock_code'].astype(str).nunique()
conn.close()
print(f'成分: {len(picked)}股 | 退市 {n_del} | 当前ST {n_st} | 有行业映射 {has_ind}/{len(picked)} (无行业={len(picked)-has_ind})')

# 直接过滤 backtest_samples 的 PIT-clean 行(复用 gen_backtest_samples 的在市/非ST/上市≥1y 过滤),
# 不做 picked×全日期 笛卡尔积(那会造 ~50% 上市前幽灵行, 特征全 NaN)
pilot = samples[samples['股票代码'].astype(str).isin(picked)].copy()
pilot['股票代码'] = pilot['股票代码'].astype(str)
pilot['报价日'] = pilot['报价日'].astype(str)
pilot.to_parquet(out, index=False)
print(f'✅ {out}: {len(picked)}股 × PIT-clean = {len(pilot)}行 ({len(pilot)/len(picked):.0f} 月均/股)')
