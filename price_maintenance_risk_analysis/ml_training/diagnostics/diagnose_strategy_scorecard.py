
import os, pickle, sys
import pandas as pd
import numpy as np
_HERE = os.path.dirname(os.path.abspath(__file__)); _ML = os.path.dirname(_HERE)
sys.path.insert(0, os.path.join(_ML, 'pipeline'))   # 管线模块已移入 pipeline/
sys.path.insert(0, _ML); os.chdir(_ML)               # 相对 data/output 路径
from train.train_scorecard import calc_iv_all_features

DATA='data/features_derived.parquet'
OLD_DIR='output/v_20260612_0009_scorecard_auc069' if os.path.exists('output/v_20260612_0009_scorecard_auc069') else 'output/v_20260612_0009_scorecard_12feat_auc069'
df=pd.read_parquet(DATA)
y=df['标签_盈利_-10']
valid=y.notna()
df=df.loc[valid].reset_index(drop=True)
y=y.loc[valid].astype(int).reset_index(drop=True)

year=pd.to_numeric(pd.to_numeric(df['报价日'], errors='coerce').astype('Int64').astype(str).str[:4], errors='coerce')
train=(year<=2024).values
test=(year>=2025).values

old_12=['个股PB','市场距离MA250','盈利能力_delta_1y','净资产增长','研发费用率','净利润','速动比率','净利率','营收_CAGR2','成长能力_delta_1y','资产负债率','现金利息负债比']
strategy=['三浪_retr','抵抗_corr_div_stock']

print('=== coverage per feature per split')
cov=[]
for f in old_12+strategy:
    tr=df.loc[train,f].notna().mean()
    te=df.loc[test,f].notna().mean()
    cov.append({'feature':f,'train_cov':tr,'test_cov':te})
print(pd.DataFrame(cov).sort_values('test_cov',ascending=False).to_string(index=False))

print('\n=== train IV')
df_train=df.loc[train].reset_index(drop=True)
X_train=df_train[old_12+strategy].apply(lambda s:pd.to_numeric(s, errors='coerce'))
iv_train=calc_iv_all_features(X_train, df_train['标签_盈利_-10'])
print(iv_train[iv_train['feature'].isin(strategy)].to_string(index=False))
print('\n... old12 IV train top 5:')
print(iv_train[iv_train['feature'].isin(old_12)].sort_values('iv', ascending=False).head(5).to_string(index=False))

print('\n=== test IV')
df_test=df.loc[test].reset_index(drop=True)
X_test=df_test[old_12+strategy].apply(lambda s:pd.to_numeric(s, errors='coerce'))
iv_test=calc_iv_all_features(X_test, df_test['标签_盈利_-10'])
print(iv_test[iv_test['feature'].isin(strategy)].to_string(index=False))
print('\n... old12 IV test top 5:')
print(iv_test[iv_test['feature'].isin(old_12)].sort_values('iv', ascending=False).head(5).to_string(index=False))

print('\n=== compare strategy vs old12 correlation on train')
corr=X_train.corr()
for f in strategy:
    top=corr[[f]].sort_values(f, ascending=False).head(6)
    print(f'\n{f}:')
    print(top.to_string())

print('\n=== raw percentiles of strategy features train vs test')
for f in strategy:
    tr=df.loc[train,f].dropna()
    te=df.loc[test,f].dropna()
    print(f'\n{f}:')
    print(pd.DataFrame({
        'train':tr.describe(percentiles=[0.05,0.25,0.5,0.75,0.95]),
        'test':te.describe(percentiles=[0.05,0.25,0.5,0.75,0.95])
    }).to_string())

print('\n=== strategy feature histogram splits on train')
for f in strategy:
    print(f'\n{f} train:')
    vals=pd.to_numeric(df.loc[train,f], errors='coerce').dropna()
    print(vals.describe())
    bins=pd.qcut(vals, q=5, duplicates='drop').value_counts().sort_index()
    print(bins)

print('\n=== strategy feature histogram splits on test')
for f in strategy:
    print(f'\n{f} test:')
    vals=pd.to_numeric(df.loc[test,f], errors='coerce').dropna()
    print(vals.describe())
    bins=pd.qcut(vals, q=5, duplicates='drop').value_counts().sort_index()
    print(bins)
