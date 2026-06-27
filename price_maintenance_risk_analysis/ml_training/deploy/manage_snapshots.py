#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据集快照库管理 CLI。

用法:
  python manage_snapshots.py list [--kind base|derived] [--label 7m]
  python manage_snapshots.py show <version>
  python manage_snapshots.py restore <version> [--out data/features_derived.parquet]

gitignore 后的数据文件可经此从 DB 快照还原。
"""
import argparse
import os
import sys

PKG = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # deploy/→ml_training/→PKG
for _p in (PKG, os.path.join(PKG,'ml_training'), os.path.join(PKG,'ml_training','pipeline'), os.path.join(PKG,'scripts')):
    if _p not in sys.path: sys.path.insert(0, _p)
import db_dataset_store as store

_DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')


def cmd_list(args):
    rows = store.list_snapshots(kind=args.kind, label_config=args.label)
    if not rows:
        print('(暂无快照)')
        return
    print(f"{'version':52} {'kind':8} {'label':12} {'shape':12} {'pos%':6} created_at")
    print('-' * 110)
    for r in rows:
        pos = f"{r['positive_rate']*100:.1f}" if r['positive_rate'] is not None else '-'
        print(f"{r['version']:52} {r['kind']:8} {r['label_config']:12} "
              f"{r['n_rows']}x{r['n_cols']:4} {pos:6} {r['created_at']}")


def cmd_show(args):
    info = next((r for r in store.list_snapshots() if r['version'] == args.version), None)
    if not info:
        sys.exit(f'快照不存在: {args.version}')
    for k, v in info.items():
        print(f"{k:14}: {v}")


def cmd_restore(args):
    out = args.out or os.path.join(_DATA, 'features_derived.parquet')
    path = store.restore_snapshot(args.version, out)
    print(f'✅ 已还原 {args.version} → {path}')


def main():
    ap = argparse.ArgumentParser(description='数据集快照库管理')
    sub = ap.add_subparsers(dest='cmd', required=True)
    p_ls = sub.add_parser('list', help='列出快照')
    p_ls.add_argument('--kind', choices=['base', 'derived'])
    p_ls.add_argument('--label', dest='label', help='label_config 过滤')
    p_ls.set_defaults(func=cmd_list)
    p_show = sub.add_parser('show', help='查看快照元信息')
    p_show.add_argument('version')
    p_show.set_defaults(func=cmd_show)
    p_rs = sub.add_parser('restore', help='还原快照到文件')
    p_rs.add_argument('version')
    p_rs.add_argument('--out', help='输出路径(默认 data/features_derived.parquet)')
    p_rs.set_defaults(func=cmd_restore)
    args = ap.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
