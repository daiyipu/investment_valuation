#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""模型元信息 DB 存储 —— 替代散落的 per-version meta.json 文件。

每训出一个模型版本, 把它的"用什么标签 + 用什么特征 + 选择漏斗 + 指标 + 关联数据快照"
存成 ml_model_meta 表的一行。一个版本一行, 可按 version/label_config 查询。
模型权重(lgb.txt/lr.pkl)仍在磁盘(二进制, predict 要用), 这里只存"配置/元信息"。

表结构(ml_model_meta):
  version        PK   版本名(如 v_20260616_1958_7m_gray_12feat_auc072)
  label_config        标签配置(7m_thr / 7m_gray ...)
  kind                thr / gray
  horizon             1/3/6/7/12
  gray_cfg            灰度阈值(如 -20/5); 阈值版为 NULL
  features            入模特征列表(JSON 文本)
  n_features          入模特征数
  selection           标准五步漏斗串(IV40→PSI→...→LGBM)
  sel_thresholds      选择阈值(JSON: n_iv/psi_max/corr_max/vif_max)
  metrics             指标(JSON: lgb_oot_auc/ks, lr_oot_auc/ks, band_spread)
  n_train/n_test/train_pos/test_pos
  dataset_version     关联 ml_dataset_snapshot.version(该模型吃的确切数据)
  note                说明
  created_at          入库时间

复用 utils.db_manager.ValuationDB.MYSQL_CONFIG(唯一 DB 配置)。
"""
import json
import os
import sys

import pymysql

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.db_manager import ValuationDB   # noqa: E402 (utils 在上级目录)


def _connect():
    return pymysql.connect(**ValuationDB.MYSQL_CONFIG)


def _ensure_table(conn):
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS ml_model_meta (
      version        VARCHAR(96) NOT NULL,
      label_config   VARCHAR(64),
      kind           VARCHAR(16),
      horizon        INT,
      gray_cfg       VARCHAR(32),
      features       LONGTEXT,
      n_features     INT,
      selection      VARCHAR(160),
      sel_thresholds LONGTEXT,
      metrics        LONGTEXT,
      n_train        INT, n_test INT,
      train_pos      DOUBLE, test_pos DOUBLE,
      dataset_version VARCHAR(96),
      note           VARCHAR(255),
      created_at     DATETIME NOT NULL,
      PRIMARY KEY (version),
      KEY idx_label (label_config),
      KEY idx_dataset (dataset_version)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4""")
    # 自愈: 给老表补权重/medians 列(权重也入 DB, 不再散落磁盘 version 目录)
    cur.execute("""SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
                   WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'ml_model_meta'""")
    existing = {r[0] for r in cur.fetchall()}
    for col, ddl in [('medians', 'LONGTEXT'), ('lgb_model', 'LONGTEXT'),
                     ('lr_bundle', 'LONGBLOB')]:
        if col not in existing:
            cur.execute(f'ALTER TABLE ml_model_meta ADD COLUMN {col} {ddl}')
    conn.commit()


def save_model_meta(meta):
    """upsert 一条模型元信息(meta dict 至少含 version)。返回 version。"""
    conn = _connect()
    try:
        _ensure_table(conn)
        cur = conn.cursor()
        row = {
            'version': meta['version'],
            'label_config': meta.get('label_config'),
            'kind': meta.get('kind'),
            'horizon': meta.get('horizon'),
            'gray_cfg': (f'{meta["gray_cfg"][0]}/{meta["gray_cfg"][1]}'
                         if meta.get('gray_cfg') else None),
            'features': json.dumps(meta.get('features', []), ensure_ascii=False),
            'n_features': meta.get('n_features'),
            'selection': meta.get('selection'),
            'sel_thresholds': json.dumps(meta.get('selection_thresholds', {}), ensure_ascii=False),
            'metrics': json.dumps(meta.get('metrics', {}), ensure_ascii=False),
            'medians': json.dumps(meta.get('medians', {}), ensure_ascii=False),
            'lgb_model': meta.get('lgb_model'),     # LGB booster 文本(模型权重)
            'lr_bundle': meta.get('lr_bundle'),     # LR {model,scaler,features} pickle 字节(权重)
            'n_train': meta.get('n_train'), 'n_test': meta.get('n_test'),
            'train_pos': meta.get('train_pos'), 'test_pos': meta.get('test_pos'),
            'dataset_version': meta.get('dataset_version'),
            'note': meta.get('note'),
        }
        cur.execute("""INSERT INTO ml_model_meta
          (version,label_config,kind,horizon,gray_cfg,features,n_features,selection,
           sel_thresholds,metrics,medians,lgb_model,lr_bundle,
           n_train,n_test,train_pos,test_pos,dataset_version,note,created_at)
          VALUES (%(version)s,%(label_config)s,%(kind)s,%(horizon)s,%(gray_cfg)s,%(features)s,
           %(n_features)s,%(selection)s,%(sel_thresholds)s,%(metrics)s,%(medians)s,%(lgb_model)s,%(lr_bundle)s,
           %(n_train)s,%(n_test)s,%(train_pos)s,%(test_pos)s,%(dataset_version)s,%(note)s,NOW())
          ON DUPLICATE KEY UPDATE
           features=VALUES(features), n_features=VALUES(n_features), selection=VALUES(selection),
           sel_thresholds=VALUES(sel_thresholds), metrics=VALUES(metrics), medians=VALUES(medians),
           lgb_model=VALUES(lgb_model), lr_bundle=VALUES(lr_bundle),
           n_train=VALUES(n_train), n_test=VALUES(n_test), train_pos=VALUES(train_pos),
           test_pos=VALUES(test_pos), dataset_version=VALUES(dataset_version), note=VALUES(note)""",
                    row)
        conn.commit()
        return meta['version']
    finally:
        conn.close()


def _row_to_meta(description, r):
    """按 cursor.description 的列名映射(不依赖列顺序, 防 ALTER 追加列错位)。"""
    names = [c[0] for c in description]
    d = dict(zip(names, r))
    for k in ('features', 'sel_thresholds', 'metrics', 'medians'):
        if d.get(k):
            try:
                d[k] = json.loads(d[k])
            except Exception:
                pass
    return d


def load_predict_bundle(version):
    """取 predict 所需的完整包: features(list) + medians(dict) + lgb_model(文本, 可空) + lr_bundle(bytes)。
    LGB 模型: lgb_model 有值 → lgb.Booster(model_str=...) + lr_bundle={model,scaler,features}。
    评分卡(SC)模型: lgb_model 为空 → lr_bundle={kind:'scorecard', woe_bins, lr_model, ...}, predict 走 WOE 路径。"""
    m = get_model_meta(version)
    if not m:
        raise KeyError(f'ml_model_meta 无版本 {version}')
    return {'features': m.get('features', []), 'medians': m.get('medians', {}),
            'lgb_model': m.get('lgb_model'), 'lr_bundle': m.get('lr_bundle')}


def get_model_meta(version):
    """取单条模型元信息。"""
    conn = _connect()
    try:
        _ensure_table(conn)
        cur = conn.cursor()
        cur.execute('SELECT * FROM ml_model_meta WHERE version=%s', (version,))
        r = cur.fetchone()
        return _row_to_meta(cur.description, r) if r else None
    finally:
        conn.close()


def list_model_metas(label_config=None, kind=None):
    """列模型元信息(可按 label_config/kind 过滤), 按创建时间倒序。"""
    conn = _connect()
    try:
        _ensure_table(conn)
        cur = conn.cursor()
        sql = 'SELECT * FROM ml_model_meta'
        where, args = [], []
        if label_config:
            where.append('label_config=%s'); args.append(label_config)
        if kind:
            where.append('kind=%s'); args.append(kind)
        if where:
            sql += ' WHERE ' + ' AND '.join(where)
        sql += ' ORDER BY created_at DESC'
        cur.execute(sql, args)
        desc = cur.description
        return [_row_to_meta(desc, r) for r in cur.fetchall()]
    finally:
        conn.close()


if __name__ == '__main__':
    # 冒烟: 建表 + 列空
    ms = list_model_metas()
    print(f'ml_model_meta 表已就绪, 当前 {len(ms)} 条模型元信息')
