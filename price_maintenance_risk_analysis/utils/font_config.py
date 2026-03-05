# -*- coding: utf-8 -*-
"""
中文字体配置工具
用于解决 matplotlib 中文显示乱码问题
"""

import matplotlib.pyplot as plt
import platform


def setup_chinese_font():
    """
    配置 matplotlib 中文字体

    根据不同操作系统自动选择合适的中文字体
    """
    system = platform.system()

    if system == 'Darwin':  # macOS
        # macOS 系统可用中文字体
        plt.rcParams['font.sans-serif'] = [
            'PingFang SC',       # 苹方（推荐）
            'STHeiti',           # 华文黑体
            'Heiti TC',          # 黑体-繁
            'Arial Unicode MS',  # Arial Unicode
            'Microsoft YaHei',   # 微软雅黑（如果在macOS上安装）
        ]
    elif system == 'Windows':
        # Windows 系统可用中文字体
        plt.rcParams['font.sans-serif'] = [
            'Microsoft YaHei',   # 微软雅黑（推荐）
            'SimHei',            # 黑体
            'SimSun',            # 宋体
            'KaiTi',             # 楷体
            'FangSong',          # 仿宋
        ]
    else:  # Linux
        # Linux 系统可用中文字体
        plt.rcParams['font.sans-serif'] = [
            'WenQuanYi Micro Hei',   # 文泉驿微米黑
            'Noto Sans CJK SC',      # Noto Sans CJK
            'DejaVu Sans',
            'SimHei',
        ]

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


if __name__ == '__main__':
    test_chinese_font()
