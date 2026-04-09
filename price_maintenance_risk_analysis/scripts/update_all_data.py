#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一数据更新脚本

功能：
按照正确顺序依次执行所有数据更新脚本：
1. update_indices_data.py - 更新指数数据
2. update_market_data.py - 生成市场数据文件

用法：
    python update_all_data.py
"""

import os
import sys
import subprocess
from datetime import datetime

# 切换到脚本目录
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
os.chdir(PROJECT_DIR)

print("="*70)
print(" 定增分析数据统一更新工具")
print("="*70)
print(f"开始时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print()

# 定义数据更新步骤（按正确顺序）
scripts_dir = SCRIPT_DIR  # SCRIPT_DIR已经是scripts目录
update_steps = [
    ('update_indices_data.py', '更新指数数据', 300),
    ('update_market_data.py', '生成市场数据文件', 300)
]

all_success = True
failed_step = None
results = {}

for i, (script_name, description, timeout) in enumerate(update_steps, 1):
    print(f"\n[{i}/{len(update_steps)}] {description}")
    print(f"脚本：{script_name}")
    print("-" * 50)

    script_path = os.path.join(scripts_dir, script_name)

    if not os.path.exists(script_path):
        print(f"❌ 脚本不存在：{script_path}")
        all_success = False
        failed_step = (script_name, description, "脚本不存在")
        break

    try:
        # 运行更新脚本
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True,
            text=True,
            timeout=timeout
        )

        # 记录结果
        results[script_name] = {
            'description': description,
            'returncode': result.returncode,
            'success': result.returncode == 0
        }

        if result.returncode == 0:
            print(f"✅ {description}成功")
            if result.stdout.strip():
                # 显示输出的前200个字符
                output_preview = result.stdout.strip()[:200]
                print(f"   输出：{output_preview}")
                if len(result.stdout.strip()) > 200:
                    print(f"   ... (还有{len(result.stdout.strip()) - 200}字符)")
        else:
            print(f"❌ {description}失败 (返回码：{result.returncode})")
            if result.stderr.strip():
                print(f"   错误：{result.stderr.strip()}")
            all_success = False
            failed_step = (script_name, description, f"返回码{result.returncode}")
            break

    except subprocess.TimeoutExpired:
        print(f"❌ {description}超时（超过{timeout}秒）")
        all_success = False
        failed_step = (script_name, description, "超时")
        break
    except Exception as e:
        print(f"❌ {description}执行出错：{e}")
        all_success = False
        failed_step = (script_name, description, str(e))
        break

print()
print("="*70)
if all_success:
    print("✅ 所有数据更新步骤完成！")
    print(f"完成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    print("现在可以重新运行报告生成：")
    print("  cd price_maintenance_risk_analysis/scripts/generate_word_report_v2")
    print("  python main.py")
else:
    print(f"❌ 数据更新失败：{failed_step[1]} - {failed_step[2]}")
    print()
    print("请检查以下内容：")
    print("1. 脚本路径是否正确")
    print("2. 网络连接是否正常")
    print("3. API Token是否有效")
    print("4. 错误信息详情")
    print()
    print("手动更新方法：")
    for script_name, description, _ in update_steps:
        print(f"  python price_maintenance_risk_analysis/scripts/{script_name}  # {description}")
    print()
    print("或者单独运行市场数据更新（最常用）：")
    print(f"  python price_maintenance_risk_analysis/scripts/update_market_data.py  # 生成市场数据文件")
    sys.exit(1)
