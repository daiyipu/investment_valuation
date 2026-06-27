#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一数据全链路更新(一键重建: 指数→行情→定增名单→收益标签→超额/短线→因子→特征导出→衍生)。

三层架构:
  L1 原始摄入 → placement_evaluation DB: recompute_label_qfq(收益) / compute_labels(超额·短线) / fetch_factors(定增结构·筹码·资金流·SMC)
  L2 DB → features.parquet: export_features(统一发射所有族+标签)
  L3 features.parquet → features_derived.parquet: derive_features(衍生比率)

按正确顺序依次执行, 每步 subprocess + 超时, 失败即停并打印手动恢复命令。

用法:
    python update_all_data.py                  # 全链路
    python update_all_data.py --from 7         # 从第7步开始(续跑)
"""
import os
import sys
import subprocess
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))                 # = scripts/data_pipeline/
PROJECT_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))               # = price_maintenance_risk_analysis/
os.chdir(PROJECT_DIR)

DP = os.path.join('scripts', 'data_pipeline')   # 取数脚本子目录(相对 PROJECT_DIR)
ML = os.path.join('ml_training')                # ml_training 目录
# (script, subdir, args, 描述, timeout秒)。fetch_factors 需 --write 才回写。
FAST, SLOW = 600, 3600
update_steps = [
    ('update_indices_data.py', DP, [], '更新指数数据', FAST),
    ('update_market_data.py', DP, [], '生成市场数据', FAST),
    ('fetch_placements.py', DP, [], '定增名单(东方财富)', FAST),
    ('recompute_label_qfq.py', DP, [], '收益标签 return_*m (qfq)', SLOW),
    ('compute_labels.py', DP, ['excess'], '超额收益 excess_*', FAST),
    ('compute_labels.py', DP, ['shortterm'], '短线收益 return_1w/2w/4w', SLOW),
    ('fetch_factors.py', DP, ['placement', '--write'], '定增结构 em_*/pp_*', FAST),
    ('fetch_factors.py', DP, ['chip', '--write'], '筹码 chip_* (cyq_chips)', SLOW),
    ('fetch_factors.py', DP, ['capitalflow', '--write'], '资金流+北向 mf_*/nb_*', SLOW),
    ('fetch_factors.py', DP, ['smc', '--write'], '聪明钱 smc_* (日/W/M)', SLOW),
    ('export_features.py', ML, [], '导出 features.parquet(统一发射)', FAST),
    ('derive_features.py', ML, ['data/features.parquet'], '衍生 features_derived.parquet', SLOW),
]


def main():
    ap = __import__('argparse').ArgumentParser(description='定增分析数据全链路统一更新')
    ap.add_argument('--from', dest='start', type=int, default=1, help='从第几步开始(1-based, 续跑)')
    args = ap.parse_args()

    print('=' * 70)
    print(' 定增分析数据全链路统一更新')
    print('=' * 70)
    print(f'开始: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} | 从第 {args.start} 步起\n')

    all_success = True
    failed = None
    for i, (script, subdir, sargs, desc, timeout) in enumerate(update_steps, 1):
        if i < args.start:
            continue
        print(f'\n[{i}/{len(update_steps)}] {desc}')
        print(f'  $ python {os.path.join(subdir, script)} {" ".join(sargs)}')
        print('-' * 60)
        script_path = os.path.join(PROJECT_DIR, subdir, script)
        if not os.path.exists(script_path):
            print(f'❌ 脚本不存在: {script_path}'); all_success = False; failed = (desc, '脚本不存在'); break
        try:
            result = subprocess.run([sys.executable, script_path] + sargs,
                                    capture_output=True, text=True, timeout=timeout)
            if result.returncode == 0:
                tail = (result.stdout or '').strip().splitlines()[-3:]
                print('✅ ' + desc + ' 成功' + ('' if not tail else '\n   ' + '\n   '.join(tail)))
            else:
                print(f'❌ {desc} 失败 (返回码 {result.returncode})')
                if result.stderr:
                    print('   ' + result.stderr.strip().splitlines()[-1][:300])
                all_success = False; failed = (desc, f'返回码{result.returncode}'); break
        except subprocess.TimeoutExpired:
            print(f'❌ {desc} 超时(>{timeout}s)'); all_success = False; failed = (desc, '超时'); break
        except Exception as e:
            print(f'❌ {desc} 执行出错: {e}'); all_success = False; failed = (desc, str(e)); break

    print('\n' + '=' * 70)
    if all_success:
        print(f'✅ 全链路完成! {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        print('   features_derived.parquet 已重建, 可训练/预测。')
    else:
        print(f'❌ 失败于: {failed[0]} ({failed[1]})')
        print('手动续跑(从失败步):')
        for i, (script, sargs, desc, _) in enumerate(update_steps, 1):
            print(f'  [{i}] {desc}: cd price_maintenance_risk_analysis && python scripts/{script} {" ".join(sargs)}')
        sys.exit(1)


if __name__ == '__main__':
    main()
