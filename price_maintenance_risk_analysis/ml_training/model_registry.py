#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
模型版本注册中心（Model Registry）

统一管理所有训练产物的版本：LightGBM+逻辑回归（full）与评分卡（scorecard）。
真相源 = `output/model_registry.json`，predict_profitability 从这里读取"当前生产版本"。

registry 结构:
{
  "current": {"full": "v_...", "scorecard": "v_..."},
  "versions": [ {version, model_type, created_at, n_features, threshold,
                 n_samples, positive_rate, metrics, dir, files, note}, ... ]
}
"""

import os
import json

# 与 train_models / train_scorecard 一致的 output 目录
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
REGISTRY_PATH = os.path.join(OUTPUT_DIR, 'model_registry.json')

VALID_TYPES = ('full', 'scorecard')


def load_registry():
    """加载 registry；不存在则返回空骨架。"""
    if not os.path.exists(REGISTRY_PATH):
        return {'current': {}, 'versions': []}
    try:
        with open(REGISTRY_PATH, 'r', encoding='utf-8') as f:
            reg = json.load(f)
    except (json.JSONDecodeError, OSError):
        return {'current': {}, 'versions': []}
    reg.setdefault('current', {})
    reg.setdefault('versions', [])
    return reg


def save_registry(reg):
    """原子写入 registry。"""
    os.makedirs(os.path.dirname(REGISTRY_PATH), exist_ok=True)
    tmp = REGISTRY_PATH + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(reg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, REGISTRY_PATH)


def register_version(model_type, version, dir, metrics, n_features=None,
                     threshold=None, n_samples=None, positive_rate=None,
                     files=None, note='', set_current=True,
                     label_config=None, dataset_version=None):
    """注册一个新版本；默认设为该 model_type 的 current。

    Args:
        model_type: 'full' | 'scorecard'
        version:   版本名（即版本目录名，如 v_20260611_1530_330feat_auc072）
        dir:       版本目录绝对路径
        metrics:   dict，如 {'lgb_auc': 0.72, 'lr_auc': 0.65} 或 {'auc': 0.69, 'ks': 0.28}
        files:     该版本包含的文件名列表
        set_current: 是否同时设为 current（默认 True）
        label_config:   训练标签配置，如 '7m_-10' / '3m_-10' / 'gray_7m'
        dataset_version: DB 快照版本(ml_dataset_snapshot.version)，冻结该模型吃的确切数据
    """
    if model_type not in VALID_TYPES:
        raise ValueError(f'model_type 必须是 {VALID_TYPES}，得到 {model_type!r}')

    reg = load_registry()
    # 去重：同名版本覆盖
    reg['versions'] = [v for v in reg['versions'] if not (
        v.get('version') == version and v.get('model_type') == model_type)]

    entry = {
        'version': version,
        'model_type': model_type,
        'created_at': _now_str(),
        'n_features': n_features,
        'threshold': threshold,
        'n_samples': n_samples,
        'positive_rate': positive_rate,
        'metrics': metrics or {},
        'dir': os.path.basename(dir.rstrip(os.sep)) if isinstance(dir, str) else dir,
        'files': files or [],
        'note': note,
        'label_config': label_config,
        'dataset_version': dataset_version,
    }
    reg['versions'].append(entry)

    if set_current:
        reg['current'][model_type] = version

    save_registry(reg)
    return entry


def get_current(model_type):
    """返回某 model_type 的当前版本名；无则 None。"""
    if model_type not in VALID_TYPES:
        raise ValueError(f'model_type 必须是 {VALID_TYPES}')
    return load_registry()['current'].get(model_type)


def set_current(model_type, version):
    """切换/回滚某 model_type 的当前版本。版本必须已注册。"""
    if model_type not in VALID_TYPES:
        raise ValueError(f'model_type 必须是 {VALID_TYPES}')
    reg = load_registry()
    exists = any(v.get('version') == version and v.get('model_type') == model_type
                 for v in reg['versions'])
    if not exists:
        raise ValueError(f'版本 {version!r} ({model_type}) 未在 registry 中注册')
    reg['current'][model_type] = version
    save_registry(reg)
    return version


def get_version_dir(model_type, version=None):
    """返回版本目录的绝对路径。version=None 表示 current。"""
    reg = load_registry()
    if version is None:
        version = reg['current'].get(model_type)
    if not version:
        return None
    return os.path.join(OUTPUT_DIR, version)


def require_current_dir(model_type):
    """返回 current 版本目录绝对路径；无则抛清晰错误。"""
    version = get_current(model_type)
    if not version:
        raise RuntimeError(
            f'registry 中无 {model_type} 类型的当前版本。请先运行 '
            f'train_models.py（full）或 train_scorecard.py（scorecard）训练并注册。')
    d = os.path.join(OUTPUT_DIR, version)
    if not os.path.isdir(d):
        raise RuntimeError(
            f'当前 {model_type} 版本 {version} 的目录不存在: {d}。'
            f'可能已被删除，请用 manage_models.py set 切换到其它版本。')
    return d


def list_versions(model_type=None):
    """返回版本列表（按注册时间倒序）。可选按 model_type 过滤。"""
    reg = load_registry()
    vers = reg['versions']
    if model_type:
        vers = [v for v in vers if v.get('model_type') == model_type]
    return list(reversed(vers))  # 最新的在前


def find_entry(version):
    """按版本名查找 entry（跨 model_type）。"""
    for v in load_registry()['versions']:
        if v.get('version') == version:
            return v
    return None


def _now_str():
    """当前时间字符串（registry 元数据用）。"""
    from datetime import datetime
    return datetime.now().strftime('%Y-%m-%d %H:%M')


if __name__ == '__main__':
    # 直接运行时打印当前状态
    reg = load_registry()
    print(f'Registry: {REGISTRY_PATH}')
    print(f'当前版本: {reg.get("current", {})}')
    print(f'已注册版本数: {len(reg.get("versions", []))}')
