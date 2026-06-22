
import os, pickle, sys
import numpy as np
import pandas as pd
_HERE = os.path.dirname(os.path.abspath(__file__)); _ML = os.path.dirname(_HERE)
sys.path.insert(0, os.path.join(_ML, 'pipeline'))   # 管线模块已移入 pipeline/
sys.path.insert(0, _ML); os.chdir(_ML)               # predict_profitability + 相对 data/output 路径
from train_scorecard import woe_transform
from predict_profitability import _parse_interval

DATA='data/features_derived.parquet'
df=pd.read_parquet(DATA)
y=df['标签_盈利_-10']
valid=y.notna(); df=df.loc[valid].reset_index(drop=True); y=y.loc[valid].astype(int).reset_index(drop=True)
year=pd.to_numeric(pd.to_numeric(df['报价日'], errors='coerce').astype('Int64').astype(str).str[:4], errors='coerce')
train=(year<=2024).values
test=(year>=2025).values

feats=['三浪_retr','抵抗_corr_div_stock']
X_train=df.loc[train,feats].apply(lambda s:pd.to_numeric(s, errors='coerce'))
y_train=y[train]
X_test=df.loc[test,feats].apply(lambda s:pd.to_numeric(s, errors='coerce'))

X_train_woe, bins=woe_transform(X_train, y_train, feats, n_bins=5)
print('=== trained bins')
for f in feats:
    print(f'\n{f}:')
    for k,v in bins[f]['woe_map'].items():
        print(f'  {k} → {v}')

def apply_bin_and_count(X, bins):
    res={}
    for f in feats:
        vals=pd.to_numeric(X[f], errors='coerce')
        parsed=[]
        for k,w in bins[f]['woe_map'].items():
            iv=_parse_interval(k)
            if iv is not None:
                parsed.append((iv[0],iv[1],w))
        parsed.sort(key=lambda t:t[0])
        min_left=parsed[0][0]
        first,last=parsed[0][2],parsed[-1][2]
        bin_counts={k:0 for k in ['in_bin','min_or_left','max_or_right','nan']}
        for v in vals.values:
            if v!=v:
                bin_counts['nan']+=1
                continue
            chosen=None
            for l,r,w in parsed:
                if l < v <= r:
                    chosen='in_bin'; break
            if chosen is None:
                if v <= min_left:
                    chosen='min_or_left'
                else:
                    chosen='max_or_right'
            bin_counts[chosen]+=1
        res[f]=bin_counts
    return res

print('\n=== bin hit on train')
count_tr=apply_bin_and_count(X_train, bins)
for f in feats:
    print(f'\n{f}:')
    t=pd.Series(count_tr[f])
    print((t/t.sum()).to_string())

print('\n=== bin hit on test')
count_te=apply_bin_and_count(X_test, bins)
for f in feats:
    print(f'\n{f}:')
    t=pd.Series(count_te[f])
    print((t/t.sum()).to_string())

print('\n=== raw bin limits vs test percentiles')
for f in feats:
    print(f'\n{f}:')
    bounds=[]
    for k in bins[f]['woe_map'].keys():
        iv=_parse_interval(k)
        if iv is not None:
            bounds.extend(list(iv))
    bounds=np.array(sorted(set(bounds)))
    te_vals=pd.to_numeric(df.loc[test,f], errors='coerce').dropna()
    te_pct=te_vals.quantile([0.05,0.25,0.5,0.75,0.95])
    print('train bounds:', bounds)
    print('test percentiles:')
    print(te_pct.to_string())
