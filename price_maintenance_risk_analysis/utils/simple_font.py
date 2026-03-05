# -*- coding: utf-8 -*-
"""
超简单字体配置 - 适用于任何环境
"""

import matplotlib.pyplot as plt
import os


def setup_font_simple():
    """
    最简单的字体配置方法
    优先级：项目字体 > 系统字体 > 英文备用
    """

    # 1. 首先检查项目目录下是否有字体文件
    project_font = os.path.join(os.path.dirname(__file__), '../fonts/SimHei.ttf')
    if os.path.exists(project_font):
        try:
            import matplotlib.font_manager as fm
            font_prop = fm.FontProperties(fname=project_font)
            font_name = font_prop.get_name()
            plt.rcParams['font.sans-serif'] = [font_name]
            plt.rcParams['axes.unicode_minus'] = False
            print(f"✅ 使用项目字体: {font_name}")
            return font_name
        except Exception as e:
            print(f"⚠️ 项目字体加载失败: {e}")

    # 2. 尝试系统字体（按优先级）
    system_fonts = [
        # macOS
        ['PingFang SC', 'STHeiti', 'Heiti TC', 'Arial Unicode MS'],
        # Windows
        ['Microsoft YaHei', 'SimHei', 'SimSun'],
        # Linux
        ['WenQuanYi Micro Hei', 'Noto Sans CJK SC', 'DejaVu Sans'],
    ]

    for font_list in system_fonts:
        try:
            plt.rcParams['font.sans-serif'] = font_list + ['DejaVu Sans', 'Arial']
            plt.rcParams['axes.unicode_minus'] = False

            # 测试是否可以使用
            fig, ax = plt.subplots(figsize=(1, 1))
            ax.text(0.5, 0.5, '测', fontsize=12)
            plt.close(fig)

            print(f"✅ 使用系统字体: {font_list[0]}")
            return font_list[0]
        except:
            continue

    # 3. 如果都失败了，使用默认字体并提示
    plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial']
    plt.rcParams['axes.unicode_minus'] = False

    print("⚠️ 无法加载中文字体，使用默认字体")
    print("   中文将显示为方框")
    print("\n   解决方案:")
    print("   1. 运行: python test_font.py 查看可用字体")
    print("   2. 下载字体文件到 data/fonts/ 目录")

    return 'DejaVu Sans'


# 提供一个极简的配置函数
def setup():
    """极简配置函数"""
    return setup_font_simple()


if __name__ == '__main__':
    setup()
