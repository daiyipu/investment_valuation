# -*- coding: utf-8 -*-
"""
直接字体配置 - 不依赖 matplotlib 字体检测
适用于 vnpy 等虚拟环境
"""

import os
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import platform


def setup_font_directly():
    """
    直接配置中文字体，使用字体文件路径
    适用于 matplotlib 无法正确检测系统字体的情况
    """
    system = platform.system()

    # 常见中文字体文件路径
    font_paths = {
        'Darwin': [
            '/System/Library/Fonts/PingFang.ttc',  # 苹方
            '/System/Library/Fonts/STHeiti Medium.ttc',  # 华文黑体
            '/System/Library/Fonts/Hiragino Sans GB.ttc',
            '/Library/Fonts/Arial Unicode.ttf',
            '/System/Library/Supplemental/Arial Unicode.ttf',
        ],
        'Linux': [
            '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
            '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
            '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',
            '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
            '/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf',
            '/usr/local/share/fonts/wqy-microhei.ttc',
            '~/.local/share/fonts/wqy-microhei.ttc',
            '~/.fonts/wqy-microhei.ttc',
        ],
        'Windows': [
            'C:\\Windows\\Fonts\\msyh.ttc',  # 微软雅黑
            'C:\\Windows\\Fonts\\simsun.ttc',  # 宋体
            'C:\\Windows\\Fonts\\simhei.ttf',  # 黑体
        ]
    }

    # 查找第一个可用的字体文件
    font_file = None
    for font_path in font_paths.get(system, []):
        expanded_path = os.path.expanduser(font_path)
        if os.path.exists(expanded_path):
            font_file = expanded_path
            break

    if font_file:
        try:
            # 直接添加字体
            font_prop = fm.FontProperties(fname=font_file)
            font_name = font_prop.get_name()

            # 设置字体
            plt.rcParams['font.sans-serif'] = [font_name]
            plt.rcParams['axes.unicode_minus'] = False

            print(f"✅ 使用字体: {font_name}")
            print(f"   路径: {font_file}")
            return font_name
        except Exception as e:
            print(f"⚠️ 加载字体失败: {e}")

    # 如果找不到字体文件，尝试备用方案
    return setup_font_fallback()


def setup_font_fallback():
    """
    备用字体配置方案
    """
    # 尝试使用常见的字体名称
    fallback_fonts = [
        'DejaVu Sans',
        'Liberation Sans',
        'Arial',
        'sans-serif'
    ]

    plt.rcParams['font.sans-serif'] = fallback_fonts
    plt.rcParams['axes.unicode_minus'] = False

    print("⚠️ 未找到中文字体文件")
    print("   使用备用字体，中文可能无法正常显示")
    print("\n请安装中文字体:")
    if platform.system() == 'Linux':
        print("   sudo apt-get install fonts-noto-cjk fonts-wqy-microhei")
    elif platform.system() == 'Darwin':
        print("   macOS 应该已有中文字体，请检查字体目录")
    else:
        print("   请确保系统已安装中文字体")

    return fallback_fonts[0]


def find_chinese_font():
    """
    在系统中查找中文字体文件
    返回找到的字体路径列表
    """
    system = platform.system()

    # 常见字体目录
    search_dirs = []

    if system == 'Darwin':
        search_dirs = [
            '/System/Library/Fonts',
            '/Library/Fonts',
            os.path.expanduser('~/Library/Fonts'),
        ]
    elif system == 'Linux':
        search_dirs = [
            '/usr/share/fonts',
            '/usr/local/share/fonts',
            os.path.expanduser('~/.local/share/fonts'),
            os.path.expanduser('~/.fonts'),
        ]
    elif system == 'Windows':
        search_dirs = [
            'C:\\Windows\\Fonts',
        ]

    # 中文字体文件名模式
    font_patterns = [
        'PingFang', 'STHeiti', 'Hiragino',  # macOS
        'noto', 'wqy', 'droid', 'cjk',      # Linux
        'msyh', 'simsun', 'simhei',         # Windows
    ]

    found_fonts = []

    for directory in search_dirs:
        if not os.path.exists(directory):
            continue

        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.endswith(('.ttf', '.ttc', '.otf')):
                    if any(pattern.lower() in file.lower() for pattern in font_patterns):
                        found_fonts.append(os.path.join(root, file))

    return found_fonts


def list_available_fonts():
    """
    列出系统中所有可用的中文字体
    """
    print("正在搜索系统中的中文字体...")
    fonts = find_chinese_font()

    if fonts:
        print(f"\n找到 {len(fonts)} 个中文字体:")
        for font in fonts:
            print(f"  - {font}")
    else:
        print("\n⚠️ 未找到中文字体文件")
        print("\n请安装中文字体:")

        system = platform.system()
        if system == 'Linux':
            print("  Ubuntu/Debian:")
            print("    sudo apt-get install fonts-noto-cjk fonts-wqy-microhei")
            print("  CentOS/RedHat:")
            print("    sudo yum install wqy-microhei-fonts")

    return fonts


# 简化的配置函数，用于 notebook
def setup():
    """简化的配置函数"""
    return setup_font_directly()


if __name__ == '__main__':
    print("=== 中文字体配置测试 ===\n")
    setup()
    list_available_fonts()
