#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
快速字体修复 - 适用于 vnpy 环境
"""

import sys
import os

print("="*60)
print("快速字体修复工具")
print("="*60)

# 1. 检查环境
print(f"\nPython 版本: {sys.version.split()[0]}")
print(f"系统: {sys.platform}")

# 2. 复制系统字体到项目目录
print("\n=== 复制系统字体 ===")

fonts_dir = os.path.join(os.path.dirname(__file__), 'fonts')
os.makedirs(fonts_dir, exist_ok=True)

# macOS 系统字体
macos_fonts = [
    '/System/Library/Fonts/STHeiti Medium.ttc',
    '/System/Library/Fonts/PingFang.ttc',
]

copied = False
for src in macos_fonts:
    if os.path.exists(src):
        filename = os.path.basename(src)
        dest = os.path.join(fonts_dir, filename)

        if not os.path.exists(dest):
            try:
                import shutil
                shutil.copy(src, dest)
                print(f"✅ 已复制: {filename}")
                copied = True
            except Exception as e:
                print(f"❌ 复制失败: {e}")
        else:
            print(f"⏭️  已存在: {filename}")
            copied = True
        break

if not copied:
    print("⚠️ 未找到系统字体")
    print("\n请手动下载 SimHei.ttf 到 fonts/ 目录")
    sys.exit(1)

# 3. 测试字体
print("\n=== 测试字体 ===")
try:
    import matplotlib
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm

    print(f"matplotlib 版本: {matplotlib.__version__}")

    # 加载项目字体
    font_files = [f for f in os.listdir(fonts_dir) if f.endswith(('.ttf', '.ttc'))]

    for font_file in font_files:
        font_path = os.path.join(fonts_dir, font_file)
        try:
            font_prop = fm.FontProperties(fname=font_path)

            # 测试绘图
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.text(0.5, 0.5, '中文测试', fontsize=30,
                   ha='center', va='center', fontproperties=font_prop)
            ax.set_title('字体测试', fontproperties=font_prop, fontsize=16)
            ax.set_xlabel('横轴', fontproperties=font_prop)
            ax.set_ylabel('纵轴', fontproperties=font_prop)

            plt.savefig('font_test_quick.png', dpi=100, bbox_inches='tight')
            plt.close()

            print(f"\n✅ 字体测试成功: {font_file}")
            print(f"   已生成测试图: font_test_quick.png")
            print(f"\n   使用方法:")
            print(f"   from utils.font_manager import get_font_prop, init_font")
            print(f"   font_prop = init_font()")
            print(f"   ax.set_title('标题', fontproperties=font_prop)")
            break
        except Exception as e:
            print(f"❌ {font_file}: {e}")

except ImportError as e:
    print(f"❌ matplotlib 未安装: {e}")
except Exception as e:
    print(f"❌ 测试失败: {e}")

print("\n" + "="*60)
print("修复完成！")
print("="*60)
print("\n下一步:")
print("1. 打开 Jupyter Notebook")
print("2. 在 notebook 第一个代码块运行:")
print("   from utils.font_manager import init_font, get_font_prop")
print("   font_prop = init_font()")
print("3. 在绘图时添加: fontproperties=font_prop")
print("4. 例如: ax.set_title('标题', fontproperties=font_prop)")
