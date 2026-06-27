#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模型版本管理 CLI

用法:
    python ml_training/manage_models.py list [--type full|scorecard]
    python ml_training/manage_models.py current
    python ml_training/manage_models.py set <full|scorecard> <version>
    python ml_training/manage_models.py info <version>
"""

import argparse
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)
sys.path.insert(0, os.path.join(SCRIPT_DIR, 'pipeline'))   # 管线模块已移入 pipeline/

import model_registry as mr


def _fmt_metrics(metrics):
    if not metrics:
        return '-'
    return '  '.join(f'{k}={v:.3f}' if isinstance(v, float) else f'{k}={v}'
                     for k, v in metrics.items())


def cmd_list(args):
    vers = mr.list_versions(args.type)
    if not vers:
        print('（无已注册版本）')
        return
    cur = mr.load_registry()['current']
    print(f'{"":2s}{"版本":<40s} {"类型":<11s} {"特征":>5s} {"AUC":>7s} {"样本":>6s} {"创建时间":<18s}')
    print('-' * 100)
    for v in vers:
        is_cur = '★' if cur.get(v['model_type']) == v['version'] else ' '
        auc = v.get('metrics', {}).get('lgb_auc') or v.get('metrics', {}).get('auc') or 0
        nfeat = v.get('n_features') or '-'
        nsamp = v.get('n_samples') or '-'
        print(f'{is_cur} {v["version"]:<39s} {v["model_type"]:<11s} '
              f'{str(nfeat):>5s} {auc:>7.3f} {str(nsamp):>6s} '
              f'{v.get("created_at", ""):<18s}')
    print('\n★ = 当前生产版本')


def cmd_current(args):
    reg = mr.load_registry()
    cur = reg.get('current', {})
    if not cur:
        print('（尚未设置任何当前版本）')
        return
    for mt in ('full', 'scorecard'):
        ver = cur.get(mt)
        if not ver:
            print(f'  {mt:11s}: (未设置)')
            continue
        entry = mr.find_entry(ver) or {}
        d = mr.get_version_dir(mt, ver)
        exists = '✓' if d and os.path.isdir(d) else '✗(目录缺失)'
        print(f'  {mt:11s}: {ver}  [{exists}]')
        if entry:
            print(f'              指标: {_fmt_metrics(entry.get("metrics"))}'
                  f'  特征={entry.get("n_features")}  样本={entry.get("n_samples")}')


def cmd_set(args):
    try:
        mr.set_current(args.type, args.version)
        print(f'✅ 已将 {args.type} 当前版本切换为: {args.version}')
    except ValueError as e:
        print(f'❌ {e}', file=sys.stderr)
        sys.exit(1)


def cmd_info(args):
    entry = mr.find_entry(args.version)
    if not entry:
        print(f'❌ 版本 {args.version} 未在 registry 中找到', file=sys.stderr)
        sys.exit(1)
    print(f'版本:     {entry["version"]}')
    print(f'类型:     {entry["model_type"]}')
    print(f'创建时间: {entry.get("created_at")}')
    print(f'特征数:   {entry.get("n_features")}')
    print(f'阈值:     {entry.get("threshold")}')
    print(f'样本数:   {entry.get("n_samples")} (盈利占比 {entry.get("positive_rate")})')
    print(f'指标:     {_fmt_metrics(entry.get("metrics"))}')
    print(f'目录:     {mr.get_version_dir(entry["model_type"], entry["version"])}')
    print(f'文件:     {", ".join(entry.get("files", [])) or "-"}')
    if entry.get('note'):
        print(f'备注:     {entry["note"]}')


def main():
    parser = argparse.ArgumentParser(description='模型版本管理 CLI')
    sub = parser.add_subparsers(dest='cmd', required=True)

    p_list = sub.add_parser('list', help='列出所有版本')
    p_list.add_argument('--type', choices=mr.VALID_TYPES, default=None)
    p_list.set_defaults(func=cmd_list)

    p_cur = sub.add_parser('current', help='显示当前生产版本')
    p_cur.set_defaults(func=cmd_current)

    p_set = sub.add_parser('set', help='切换/回滚当前版本')
    p_set.add_argument('type', choices=mr.VALID_TYPES, help='模型类型')
    p_set.add_argument('version', help='目标版本名（见 list）')
    p_set.set_defaults(func=cmd_set)

    p_info = sub.add_parser('info', help='查看某版本详情')
    p_info.add_argument('version')
    p_info.set_defaults(func=cmd_info)

    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
