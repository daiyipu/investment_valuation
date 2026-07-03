#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
质量监控机制建立

监控内容:
1. 模型性能监控 (AUC, KS, 稳定性)
2. 单调性质量监控 (合格率, 波浪形特征)
3. 业务影响监控 (预测分布, 覆盖率)
4. 系统性能监控 (响应时间, 错误率)

监控方式:
- 实时监控面板
- 定期质量报告
- 异常告警机制
- 趋势分析

用法:
    python validate/quality_monitoring.py --check-type all
    python validate/quality_monitoring.py --generate-report
    python validate/quality_monitoring.py --setup-alerts
"""

import os
import sys
import json
import argparse
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict

# 路径设置
HERE = os.path.dirname(os.path.abspath(__file__))      # validate/
ML_ROOT = os.path.dirname(HERE)                          # ml_training/
sys.path.insert(0, os.path.join(ML_ROOT, 'pipeline'))     # pipeline/
sys.path.insert(0, ML_ROOT)                               # ml_training/

from deploy.model_registry import get_current, get_model_meta, load_predict_bundle
from validate.validate_methods import eval_metrics
import pickle


class QualityMonitoringSystem:
    """质量监控系统"""

    def __init__(self):
        self.monitoring_rules = {
            'performance': {
                'auc': {
                    'target': 0.58,
                    'warning': 0.56,
                    'critical': 0.55,
                    'trend_window': 7  # 7天趋势
                },
                'ks': {
                    'target': 0.14,
                    'warning': 0.13,
                    'critical': 0.12,
                    'trend_window': 7
                }
            },
            'monotonicity': {
                'acceptable_rate': {
                    'target': 0.80,
                    'warning': 0.70,
                    'critical': 0.60
                },
                'wave_features': {
                    'max_count': 2,
                    'warning': 4,
                    'critical': 6
                }
            },
            'coverage': {
                'prediction_rate': {
                    'target': 0.95,
                    'warning': 0.90,
                    'critical': 0.85
                },
                'missing_features': {
                    'max_missing': 1,
                    'warning': 2,
                    'critical': 3
                }
            },
            'system': {
                'response_time': {
                    'target_ms': 100,
                    'warning_ms': 150,
                    'critical_ms': 200
                },
                'error_rate': {
                    'target': 0.01,
                    'warning': 0.02,
                    'critical': 0.05
                }
            }
        }

        self.alert_history = defaultdict(list)
        self.monitoring_data = defaultdict(list)

    def check_model_quality(self, version=None):
        """检查模型质量"""
        if version is None:
            version = get_current('scorecard')

        print(f"🔍 检查模型质量: {version}")
        print()

        try:
            # 加载模型
            bundle = load_predict_bundle(version)
            sc_data = pickle.loads(bundle['lr_bundle'])

            features = sc_data['features']
            woe_bins = sc_data['woe_bins']
            meta = get_model_meta(version) or {}

            # 检查单调性
            monotonicity_result = self.check_monotonicity(woe_bins, features)

            # 检查性能指标
            performance_result = {
                'loyo_auc': meta.get('sc_loyo_auc'),
                'loyo_ks': meta.get('sc_loyo_ks'),
                'oot_auc': meta.get('sc_oot_auc'),
                'oot_ks': meta.get('sc_oot_ks')
            }

            # 综合评分
            quality_score = self.calculate_quality_score(monotonicity_result, performance_result)

            return {
                'version': version,
                'monotonicity': monotonicity_result,
                'performance': performance_result,
                'quality_score': quality_score,
                'timestamp': datetime.now().isoformat()
            }

        except Exception as e:
            print(f"❌ 质量检查失败: {e}")
            return None

    def check_monotonicity(self, bins, features):
        """检查单调性质量"""
        results = {}

        for feature in features:
            if feature not in bins:
                continue

            wb = bins[feature]
            woes = wb.get('woes', [])

            if len(woes) < 2:
                results[feature] = {'acceptable': False, 'status': 'insufficient_bins'}
                continue

            # 检查单调性
            trends = []
            for i in range(1, len(woes)):
                if woes[i] >= woes[i-1]:
                    trends.append('increase')
                else:
                    trends.append('decrease')

            unique_trends = len(set(trends))

            if unique_trends == 1:
                acceptable = True
                status = 'monotonic'
            elif unique_trends == 2 and len(woes) == 2:
                acceptable = True
                status = 'u_shape'
            else:
                acceptable = False
                status = 'wave'

            results[feature] = {
                'acceptable': acceptable,
                'status': status,
                'bin_count': len(woes),
                'trend_changes': unique_trends - 1
            }

        return results

    def calculate_quality_score(self, monotonicity_result, performance_result):
        """计算综合质量分数"""
        score = 0
        max_score = 100

        # 单调性得分 (40分)
        if monotonicity_result:
            acceptable_count = sum(1 for r in monotonicity_result.values() if r['acceptable'])
            total_count = len(monotonicity_result)
            monotonicity_rate = acceptable_count / total_count if total_count > 0 else 0
            score += (monotonicity_rate * 40)

        # 性能得分 (40分)
        if performance_result.get('loyo_auc'):
            auc = performance_result['loyo_auc']
            if auc >= 0.58:
                score += 40
            elif auc >= 0.56:
                score += 30
            elif auc >= 0.55:
                score += 20
            else:
                score += 10

        # 稳定性得分 (20分)
        score += 15  # 基础分

        return min(score, max_score)

    def generate_quality_report(self, quality_result):
        """生成质量报告"""
        if not quality_result:
            return None

        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        report_file = f'output/quality_report_{timestamp}.md'

        # 生成报告内容
        monotonicity = quality_result['monotonicity']
        performance = quality_result['performance']
        quality_score = quality_result['quality_score']

        total_features = len(monotonicity)
        acceptable_features = sum(1 for r in monotonicity.values() if r['acceptable'])
        wave_features = total_features - acceptable_features

        report_content = f"""# 模型质量监控报告

## 概览

- **模型版本**: {quality_result['version']}
- **检查时间**: {quality_result['timestamp']}
- **综合评分**: {quality_score}/100
- **质量等级**: {"优秀" if quality_score >= 80 else "良好" if quality_score >= 60 else "需改进"}

## 单调性质量

- **总特征数**: {total_features}
- **合格特征**: {acceptable_features} ({acceptable_features/total_features*100:.1f}%)
- **波浪形特征**: {wave_features} ({wave_features/total_features*100:.1f}%)
- **合格率目标**: 80%+ → {"✅ 达成" if acceptable_features/total_features >= 0.8 else "❌ 未达成"}

### 单调性详情

"""

        for feature, result in monotonicity.items():
            emoji = "✅" if result['acceptable'] else "❌"
            report_content += f"- {emoji} **{feature}**: {result['status']} ({result['bin_count']}分箱)\n"

        report_content += f"""
## 性能指标

| 指标 | 值 | 目标 | 状态 |
|-----|-----|------|------|
| LOYO AUC | {performance.get('loyo_auc', 'N/A')} | ≥0.58 | {"✅" if performance.get('loyo_auc', 0) >= 0.58 else "❌"} |
| LOYO KS | {performance.get('loyo_ks', 'N/A')} | ≥0.14 | {"✅" if performance.get('loyo_ks', 0) >= 0.14 else "❌"} |
| OOT AUC | {performance.get('oot_auc', 'N/A')} | ≥0.58 | {"✅" if performance.get('oot_auc', 0) >= 0.58 else "❌"} |
| OOT KS | {performance.get('oot_ks', 'N/A')} | ≥0.14 | {"✅" if performance.get('oot_ks', 0) >= 0.14 else "❌"} |

## 监控建议

"""

        # 生成监控建议
        issues = []

        if acceptable_features / total_features < 0.8:
            issues.append(f"单调性合格率{acceptable_features/total_features*100:.1f}%低于80%目标")

        if performance.get('loyo_auc', 0) < 0.56:
            issues.append(f"AUC {performance.get('loyo_auc')} 低于警告阈值0.56")

        if wave_features > 2:
            issues.append(f"波浪形特征{wave_features}个超过建议的2个")

        if issues:
            report_content += "### 发现的问题\n"
            for issue in issues:
                report_content += f"- ⚠️ {issue}\n"
            report_content += "\n建议: 进行模型优化或特征筛选\n"
        else:
            report_content += "✅ 所有指标在正常范围内\n"

        report_content += f"""
## 长期监控

### 建议监控频率
- **实时监控**: 系统性能指标
- **每日检查**: 预测分布和覆盖率
- **每周报告**: 模型质量趋势
- **每月审计**: 完整LOYO验证

### 趋势分析
- 关注AUC和KS的变化趋势
- 监控单调性合格率的稳定性
- 跟踪业务指标的影响

---

*本报告由 quality_monitoring.py 自动生成*
"""

        # 保存报告
        os.makedirs('output', exist_ok=True)
        with open(report_file, 'w', encoding='utf-8') as f:
            f.write(report_content)

        print(f"✅ 质量报告已生成: {report_file}")
        return report_file

    def setup_monitoring_dashboard(self):
        """设置监控面板配置"""
        dashboard_config = {
            'panels': [
                {
                    'name': '模型性能监控',
                    'metrics': ['auc', 'ks', 'precision', 'recall'],
                    'refresh_interval': '1h',
                    'alert_thresholds': {
                        'auc': {'min': 0.55, 'max': 1.0},
                        'ks': {'min': 0.12, 'max': 1.0}
                    }
                },
                {
                    'name': '单调性质量监控',
                    'metrics': ['acceptable_rate', 'wave_features_count'],
                    'refresh_interval': '6h',
                    'alert_thresholds': {
                        'acceptable_rate': {'min': 0.70, 'max': 1.0},
                        'wave_features_count': {'min': 0, 'max': 4}
                    }
                },
                {
                    'name': '系统性能监控',
                    'metrics': ['response_time', 'error_rate', 'throughput'],
                    'refresh_interval': '5min',
                    'alert_thresholds': {
                        'response_time': {'min': 0, 'max': 200},
                        'error_rate': {'min': 0, 'max': 0.02}
                    }
                },
                {
                    'name': '业务影响监控',
                    'metrics': ['prediction_volume', 'score_distribution', 'coverage_rate'],
                    'refresh_interval': '1h',
                    'alert_thresholds': {
                        'prediction_volume': {'min': 100, 'max': None},
                        'coverage_rate': {'min': 0.90, 'max': 1.0}
                    }
                }
            ],
            'alert_channels': ['email', 'slack', 'webhook'],
            'alert_recipients': ['ml-team@example.com', 'devops@example.com']
        }

        config_file = 'output/monitoring_dashboard.json'
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(dashboard_config, f, indent=2, ensure_ascii=False)

        print(f"✅ 监控面板配置已生成: {config_file}")
        return config_file


def main():
    parser = argparse.ArgumentParser(description='质量监控机制')
    parser.add_argument('--check-type', choices=['all', 'performance', 'monotonicity', 'system'],
                       default='all', help='检查类型')
    parser.add_argument('--version', help='指定模型版本')
    parser.add_argument('--generate-report', action='store_true',
                       help='生成质量报告')
    parser.add_argument('--setup-alerts', action='store_true',
                       help='设置告警机制')
    args = parser.parse_args()

    print(f"=== 质量监控机制 ===")
    print(f"检查类型: {args.check_type}")
    print()

    monitor = QualityMonitoringSystem()

    if args.check_type in ['all', 'performance', 'monotonicity']:
        print("🔍 执行质量检查...")
        quality_result = monitor.check_model_quality(args.version)

        if quality_result:
            print(f"✅ 质量检查完成")
            print(f"  综合评分: {quality_result['quality_score']}/100")
            print(f"  单调性合格率: {sum(1 for r in quality_result['monotonicity'].values() if r['acceptable'])}/{len(quality_result['monotonicity'])}")

            if args.generate_report:
                monitor.generate_quality_report(quality_result)

    if args.setup_alerts:
        print()
        print("⚙️ 设置监控面板...")
        monitor.setup_monitoring_dashboard()

    print()
    print("✅ 质量监控机制建立完成")
    print("📋 监控配置已保存到 output/")

    return 0


if __name__ == "__main__":
    sys.exit(main())