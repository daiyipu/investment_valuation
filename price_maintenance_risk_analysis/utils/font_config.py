# -*- coding: utf-8 -*-
"""
中文字体配置工具
用于解决 matplotlib 中文显示乱码问题
支持 vnpy 等特殊环境
"""

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import platform
import os
import warnings


def get_available_chinese_fonts():
    """
    获取系统中可用的中文字体列表
    """
    available_fonts = []
    font_keywords = [
        'ping', 'hei', 'song', 'kai', 'yuan', 'noto', 'cjk', 'han',
        'chinese', 'yahei', 'simhei', 'simsun', 'st', 'kaiti', 'fang',
        'source', 'wenquanyi', 'wqy', 'droid', 'ukai', 'uming'
    ]

    try:
        for font in fm.fontManager.ttflist:
            font_name = font.name
            # 检查是否包含中文字体关键词
            if any(keyword in font_name.lower() for keyword in font_keywords):
                if font_name not in available_fonts:
                    available_fonts.append(font_name)
    except Exception as e:
        warnings.warn(f"获取字体列表时出错: {e}")

    return available_fonts


def setup_chinese_font():
    """
    配置 matplotlib 中文字体

    根据不同操作系统自动选择合适的中文字体
    优先使用系统中实际可用的字体
    """
    system = platform.system()

    # 获取系统中实际可用的中文字体
    available_fonts = get_available_chinese_fonts()

    # 定义候选字体列表（按优先级排序）
    if system == 'Darwin':  # macOS
        font_candidates = [
            'PingFang SC', 'PingFang-SC', 'PingFang',
            'STHeiti', 'STHeitiSC', 'Heiti SC', 'Heiti',
            'Arial Unicode MS',
            'Hiragino Sans GB',
            'Microsoft YaHei',  # 如果安装了Office
        ]
    elif system == 'Windows':
        font_candidates = [
            'Microsoft YaHei', 'Microsoft YaHei UI',
            'SimHei',
            'SimSun',
            'KaiTi', 'KaiTi_GB2312',
            'FangSong', 'FangSong_GB2312',
        ]
    else:  # Linux / vnpy环境
        font_candidates = [
            'WenQuanYi Micro Hei', 'WenQuanYi Micro Hei Mono',
            'Noto Sans CJK SC', 'Noto Sans CJK',
            'Wqy Micro Hei', 'Wqy-Micro-Hei',
            'Droid Sans Fallback',
            'Source Han Sans CN', 'Source Han Sans',
            'AR PL UMing CN', 'AR PL UKai CN',
            'SimHei',
            'DejaVu Sans',
        ]

    # 优先使用系统中实际可用的字体
    selected_fonts = []
    for candidate in font_candidates:
        if candidate in available_fonts or any(candidate.lower() in f.lower() for f in available_fonts):
            selected_fonts.append(candidate)

    # 如果没有找到任何字体，添加所有候选字体
    if not selected_fonts:
        selected_fonts = font_candidates

    # 如果系统中检测到其他中文字体，也添加进来
    for font in available_fonts:
        if font not in selected_fonts and not any(f.lower() in font.lower() for f in selected_fonts):
            selected_fonts.append(font)

    # 设置字体
    plt.rcParams['font.sans-serif'] = selected_fonts

    # 解决负号显示问题
    plt.rcParams['axes.unicode_minus'] = False

    # 设置字体大小
    plt.rcParams['font.size'] = 10

    return plt.rcParams['font.sans-serif']


def test_chinese_font():
    """
    测试中文字体显示效果
    """
    setup_chinese_font()

    import matplotlib.pyplot as plt
    import numpy as np

    fig, ax = plt.subplots(figsize=(10, 6))

    x = np.linspace(0, 10, 100)
    ax.plot(x, np.sin(x), label='正弦波')
    ax.plot(x, np.cos(x), label='余弦波')

    ax.set_xlabel('横轴标签')
    ax.set_ylabel('纵轴标签')
    ax.set_title('中文字体测试 - 中文显示正常')
    ax.legend()
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig('font_test.png', dpi=100, bbox_inches='tight')
    plt.show()

    print("✅ 中文字体测试完成，请查看输出的图片")
    print(f"当前使用字体: {plt.rcParams['font.sans-serif'][0]}")


def diagnose_font_issue():
    """
    诊断字体配置问题并提供解决方案
    """
    print("="*60)
    print("中文字体诊断")
    print("="*60)

    # 1. 检查系统
    system = platform.system()
    print(f"\n操作系统: {system}")

    # 2. 检查可用字体
    available_fonts = get_available_chinese_fonts()
    print(f"\n检测到的中文字体数量: {len(available_fonts)}")

    if available_fonts:
        print(f"可用字体（前10个）:")
        for font in available_fonts[:10]:
            print(f"  - {font}")
    else:
        print("⚠️ 未检测到中文字体！")

    # 3. 配置字体
    selected = setup_chinese_font()
    print(f"\n已配置字体（前5个）:")
    for font in selected[:5]:
        print(f"  - {font}")

    # 4. 提供安装建议
    if not available_fonts:
        print("\n" + "="*60)
        print("安装中文字体建议:")
        print("="*60)

        if system == 'Darwin':
            print("\nmacOS 系统通常已内置中文字体。")
            print("如果仍有问题，请尝试:")
            print("  1. 打开'字体册'应用，查找'苹方'或'华文黑体'")
            print("  2. 确保字体已启用")
        elif system == 'Windows':
            print("\nWindows 系统请安装中文字体:")
            print("  1. 下载微软雅黑字体包")
            print("  2. 解压后右键选择'安装'")
        else:
            print("\nLinux/vnpy 环境请安装中文字体:")
            print("  Ubuntu/Debian:")
            print("    sudo apt-get install fonts-wqy-microhei fonts-wqy-zenhei")
            print("    sudo apt-get install fonts-noto-cjk")
            print("\n  CentOS/RedHat:")
            print("    sudo yum install wqy-microhei-fonts wqy-zenhei-fonts")
            print("\n  或下载字体文件到 ~/.fonts/ 目录:")
            print("    mkdir -p ~/.fonts")
            print("    # 将字体文件复制到该目录")
            print("    fc-cache -fv")

    # 5. 测试
    print("\n" + "="*60)
    print("运行测试...")
    print("="*60)
    test_chinese_font()

    return {
        'system': system,
        'available_fonts': available_fonts,
        'selected_fonts': selected
    }


def clear_font_cache():
    """
    清除 matplotlib 字体缓存
    当新安装字体后需要运行此函数
    """
    import shutil

    try:
        # 获取缓存目录
        cache_dir = fm.get_cachedir()
        print(f"字体缓存目录: {cache_dir}")

        if os.path.exists(cache_dir):
            # 删除缓存
            shutil.rmtree(cache_dir)
            print("✅ 字体缓存已清除")
            print("⚠️ 请重启 Python/Jupyter 后重新导入 matplotlib")
        else:
            print("缓存目录不存在")

    except Exception as e:
        print(f"清除缓存时出错: {e}")


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == 'diagnose':
        diagnose_font_issue()
    elif len(sys.argv) > 1 and sys.argv[1] == 'clear-cache':
        clear_font_cache()
    else:
        test_chinese_font()


if __name__ == '__main__':
    test_chinese_font()
