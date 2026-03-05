#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
快速字体测试脚本
在 vnpy 环境中运行此脚本诊断字体问题
"""

import sys
import os
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import platform

print("="*60)
print("字体诊断工具")
print("="*60)

# 1. 基本信息
print(f"\nPython 版本: {sys.version}")
print(f"matplotlib 版本: {matplotlib.__version__}")
print(f"系统: {platform.system()}")

# 2. matplotlib 缓存位置
print(f"\nmatplotlib 缓存目录:")
try:
    cache_dir = fm.get_cachedir()
    print(f"  {cache_dir}")
except AttributeError:
    # 旧版本 matplotlib 兼容
    import tempfile
    cache_dir = os.path.join(tempfile.gettempdir(), 'matplotlib-cache')
    print(f"  {cache_dir} (默认位置)")

# 3. 列出所有可用字体
print(f"\n=== 所有字体 ===")
try:
    all_fonts = [f.name for f in fm.fontManager.ttflist]
    print(f"总字体数: {len(all_fonts)}")

    # 查找中文字体
    chinese_keywords = ['ping', 'hei', 'song', 'kai', 'noto', 'wenquanyi', 'wqy', 'cjk', 'han', 'chinese', 'yahei', 'simhei', 'simsun', 'st', 'kaiti', 'fang']
    chinese_fonts = []
    for font in all_fonts:
        if any(kw in font.lower() for kw in chinese_keywords):
            chinese_fonts.append(font)

    chinese_fonts = sorted(list(set(chinese_fonts)))

    print(f"\n找到 {len(chinese_fonts)} 个中文字体:")
    for font in chinese_fonts:
        print(f"  - {font}")

    if not chinese_fonts:
        print("\n⚠️ 没有找到中文字体！")
        print("\n解决方案:")
        print("  1. 安装中文字体包")
        print("  2. 下载字体文件到项目目录")

except Exception as e:
    print(f"获取字体列表时出错: {e}")

# 4. 尝试绘制测试图
print(f"\n=== 测试绘图 ===")
try:
    # 尝试使用不同的字体
    test_fonts = ['sans-serif']
    if chinese_fonts:
        test_fonts = chinese_fonts[:3]

    for font_name in test_fonts:
        try:
            plt.rcParams['font.sans-serif'] = [font_name]
            plt.rcParams['axes.unicode_minus'] = False

            fig, ax = plt.subplots(figsize=(8, 4))
            ax.text(0.5, 0.5, '测试中文显示', fontsize=20, ha='center', va='center')
            ax.set_title(f'字体: {font_name}')
            ax.set_xlabel('横轴标签')
            ax.set_ylabel('纵轴标签')

            filename = f'font_test_{font_name.replace(" ", "_")}.png'
            plt.savefig(filename, dpi=80, bbox_inches='tight')
            plt.close()

            print(f"  ✅ {font_name}: 已生成测试图 {filename}")

        except Exception as e:
            print(f"  ❌ {font_name}: 失败 - {e}")

except Exception as e:
    print(f"测试绘图时出错: {e}")

# 5. 检查系统字体目录
print(f"\n=== 系统字体目录检查 ===")
system = platform.system()
font_dirs = []

if system == 'Darwin':  # macOS
    font_dirs = [
        '/System/Library/Fonts',
        '/Library/Fonts',
        os.path.expanduser('~/Library/Fonts'),
    ]
elif system == 'Linux':
    font_dirs = [
        '/usr/share/fonts',
        '/usr/local/share/fonts',
        os.path.expanduser('~/.local/share/fonts'),
        os.path.expanduser('~/.fonts'),
    ]
elif system == 'Windows':
    font_dirs = ['C:\\Windows\\Fonts']

for font_dir in font_dirs:
    exists = os.path.exists(font_dir)
    print(f"  {'✅' if exists else '❌'} {font_dir}")

    if exists:
        # 列出一些字体文件
        try:
            files = os.listdir(font_dir)[:10]
            for f in files:
                if f.endswith(('.ttf', '.ttc', '.otf')):
                    print(f"      - {f}")
        except:
            pass

print("\n" + "="*60)
print("诊断完成")
print("="*60)
print("\n如果中文仍然显示乱码，请尝试:")
print("  1. 安装系统字体包")
print("  2. 清除 matplotlib 缓存: python -c \"import matplotlib.font_manager; matplotlib.font_manager._rebuild()\"")
print("  3. 或者使用下面的方案下载字体到项目目录")
