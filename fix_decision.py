#!/usr/bin/env python3
"""从已有子场景数据重新计算定增决策（不需重跑分析）。

用法: python fix_decision.py <scored_excel_path>
"""
import sys
import pandas as pd

def fix(filepath):
    df = pd.read_excel(filepath, sheet_name='Sheet1')
    print(f'读取 {len(df)} 条')

    sub_map = {
        'step1': ['市场指数', '行业PE', '个股PE'],
        'step2': ['DCF估值', '修正PE估值'],
        'step3': ['参数构造', '蒙特卡洛', '反向推算'],
    }

    pass_count = 0
    for idx, row in df.iterrows():
        # 各步通过判断
        results = {}
        for step, cols in sub_map.items():
            passed = sum(1 for c in cols if c in row and '✓' in str(row[c]))
            if step == 'step1':
                results['step1'] = passed >= 2
            elif step == 'step2':
                results['step2'] = passed >= 1
            elif step == 'step3':
                results['step3'] = passed >= 2

        all_pass = results['step1'] and results['step2'] and results['step3']
        decision = '建议参与本次定向增发' if all_pass else '不建议本阶段参与该企业的本笔定向增发'
        df.at[idx, '定增决策'] = decision
        if all_pass:
            pass_count += 1

    print(f'修复后通过: {pass_count}/{len(df)}')

    # 同时更新最终结论
    if '综合趋势' in df.columns and '最终结论' in df.columns:
        df['最终结论'] = df.apply(
            lambda r: '通过' if r['定增决策'] == '建议参与本次定向增发' and str(r.get('综合趋势', '')) == '通过' else '不通过',
            axis=1
        )
        final_pass = (df['最终结论'] == '通过').sum()
        print(f'最终结论通过: {final_pass}/{len(df)}')

    df.to_excel(filepath, index=False, sheet_name='Sheet1', engine='openpyxl')
    print(f'✅ 已更新: {filepath}')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python fix_decision.py <scored_excel_path>')
        sys.exit(1)
    fix(sys.argv[1])
