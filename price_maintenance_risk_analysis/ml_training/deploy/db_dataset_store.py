#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据集版本快照库（ml_dataset_snapshot）。

把 export_features/derive_features 产出的冻结特征矩阵按版本存入 MySQL，
使 gitignore 后的数据文件仍可按模型版本精确还原——qfq 会漂移（今天重建 ≠
以后精确复现），故把「某模型当年到底吃了什么数据」冻结成 parquet 字节入库。

存储 = parquet 字节进 LONGBLOB，整存整取（快照从不切片查询），按 sha256 去重。
DB 配置复用 utils.db_manager.ValuationDB.MYSQL_CONFIG（唯一定义点，不另复制）。
"""
import io
import os
import sys
import hashlib
from datetime import datetime

import pandas as pd
import pymysql

# 复用唯一 DB 配置（utils 在父目录）
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.db_manager import ValuationDB

_CFG = ValuationDB.MYSQL_CONFIG
_TABLE = 'ml_dataset_snapshot'

_DDL = f"""CREATE TABLE IF NOT EXISTS {_TABLE} (
  version       VARCHAR(96)  NOT NULL,
  kind          ENUM('base','derived') NOT NULL,
  label_config  VARCHAR(64)  NOT NULL,
  created_at    DATETIME     NOT NULL,
  n_rows        INT          NOT NULL,
  n_cols        INT          NOT NULL,
  positive_rate DOUBLE       NULL,
  sha256        CHAR(64)     NOT NULL,
  parquet       LONGBLOB     NOT NULL,
  note          VARCHAR(255) NULL,
  PRIMARY KEY (version),
  KEY idx_sha (sha256),
  KEY idx_kind_label (kind, label_config)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4"""


def _connect():
    """普通游标连接（非 DictCursor 适配器，BLOB 原样 bytes 往返）。"""
    return pymysql.connect(**_CFG)


def _ensure_table(conn):
    with conn.cursor() as cur:
        cur.execute(_DDL)
    conn.commit()


def save_snapshot(df, kind, label_config, note=''):
    """把 df 存为快照；相同 sha256 复用已存版本，不重复存。返回 version 标签。

    Args:
        df:           特征 DataFrame
        kind:         'base' | 'derived'
        label_config: 标签集标签，如 '7m' / '1m_-10' / 'gray_7m'
        note:         备注（lineage 等）
    Returns:
        version 字符串（自描述: <kind>_<yyyymmdd_HHMM>_<sha8>_<label_config>）
    """
    buf = io.BytesIO()
    df.to_parquet(buf, index=False)
    data = buf.getvalue()
    sha = hashlib.sha256(data).hexdigest()
    sha8 = sha[:8]

    conn = _connect()
    try:
        _ensure_table(conn)
        with conn.cursor() as cur:
            cur.execute(f"SELECT version FROM {_TABLE} WHERE sha256=%s", (sha,))
            row = cur.fetchone()
            if row:
                return row[0]  # 已存，复用同一 version

        # positive_rate：取 标签_盈利_-10 均值作信息列（基线口径，缺则 NULL）
        pos = None
        if '标签_盈利_-10' in df.columns:
            pos = float(pd.to_numeric(df['标签_盈利_-10'], errors='coerce').mean())

        version = f"{kind}_{datetime.now().strftime('%Y%m%d_%H%M')}_{sha8}_{label_config}"
        with conn.cursor() as cur:
            cur.execute(
                f"INSERT INTO {_TABLE} (version,kind,label_config,created_at,"
                f"n_rows,n_cols,positive_rate,sha256,parquet,note) "
                f"VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                (version, kind, label_config, datetime.now(),
                 int(len(df)), int(len(df.columns)), pos, sha, data, note[:255]))
        conn.commit()
        return version
    finally:
        conn.close()


def load_snapshot(version):
    """按 version 还原 DataFrame（逐列与原 parquet 完全一致）。"""
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(f"SELECT parquet FROM {_TABLE} WHERE version=%s", (version,))
            row = cur.fetchone()
        if not row:
            raise KeyError(f'快照不存在: {version}')
        return pd.read_parquet(io.BytesIO(row[0]))
    finally:
        conn.close()


def list_snapshots(kind=None, label_config=None):
    """返回快照元信息列表（不含 BLOB），按 created_at 倒序。"""
    conn = _connect()
    try:
        _ensure_table(conn)
        sql = (f"SELECT version,kind,label_config,created_at,n_rows,n_cols,"
               f"positive_rate,sha256,note FROM {_TABLE}")
        where, params = [], []
        if kind:
            where.append('kind=%s'); params.append(kind)
        if label_config:
            where.append('label_config=%s'); params.append(label_config)
        if where:
            sql += ' WHERE ' + ' AND '.join(where)
        sql += ' ORDER BY created_at DESC'
        with conn.cursor() as cur:
            cur.execute(sql, params)
            cols = [d[0] for d in cur.description]
            rows = cur.fetchall()
        return [dict(zip(cols, r)) for r in rows]
    finally:
        conn.close()


def restore_snapshot(version, out_path):
    """还原快照到 parquet（+同名 csv）文件路径。"""
    df = load_snapshot(version)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    df.to_parquet(out_path, index=False)
    if out_path.endswith('.parquet'):
        df.to_csv(out_path.replace('.parquet', '.csv'), index=False)
    return out_path


if __name__ == '__main__':
    rows = list_snapshots()
    if not rows:
        print('(暂无快照)')
    for r in rows:
        print(f"{r['version']}  {r['kind']:8} {r['label_config']:10} "
              f"{r['n_rows']}x{r['n_cols']}  {r['created_at']}")
