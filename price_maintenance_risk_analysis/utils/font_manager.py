# -*- coding: utf-8 -*-
"""
终极字体解决方案
直接在每个绘图元素上设置字体属性
"""

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os
import platform


# 全局字体属性对象
_font_prop = None


def get_system_font_path():
    """
    获取系统中的中文字体文件路径
    """
    system = platform.system()

    # 常见字体路径
    font_paths = {
        'Darwin': [
            '/System/Library/Fonts/PingFang.ttc',
            '/System/Library/Fonts/STHeiti Medium.ttc',
            '/System/Library/Fonts/Hiragino Sans GB.ttc',
            '/Library/Fonts/Arial Unicode.ttf',
        ],
        'Linux': [
            '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
            '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc',
            '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
        ],
        'Windows': [
            'C:\\Windows\\Fonts\\msyh.ttc',
            'C:\\Windows\\Fonts\\simsun.ttc',
            'C:\\Windows\\Fonts\\simhei.ttf',
        ]
    }

    for font_path in font_paths.get(system, []):
        if os.path.exists(font_path):
            return font_path

    return None


def init_font():
    """
    初始化字体，返回 FontProperties 对象
    """
    global _font_prop

    # 先尝试项目目录
    project_fonts = [
        '../fonts/SimHei.ttf',
        '../fonts/simhei.ttf',
        '../fonts/msyh.ttc',
    ]

    for font_path in project_fonts:
        full_path = os.path.join(os.path.dirname(__file__), font_path)
        if os.path.exists(full_path):
            try:
                _font_prop = fm.FontProperties(fname=full_path)
                print(f"✅ 使用项目字体: {font_path}")
                return _font_prop
            except Exception as e:
                print(f"⚠️ 项目字体加载失败: {e}")

    # 尝试系统字体
    system_font_path = get_system_font_path()
    if system_font_path:
        try:
            _font_prop = fm.FontProperties(fname=system_font_path)
            print(f"✅ 使用系统字体: {system_font_path}")
            return _font_prop
        except Exception as e:
            print(f"⚠️ 系统字体加载失败: {e}")

    # 如果都失败了，使用默认字体并返回 None
    print("⚠️ 无法加载中文字体，图表中文可能显示异常")
    _font_prop = None
    return None


def get_font_prop():
    """
    获取字体属性对象，用于绘图时设置字体
    使用方法:
        ax.set_title('标题', fontproperties=get_font_prop())
        ax.set_xlabel('X轴', fontproperties=get_font_prop())
    """
    global _font_prop
    if _font_prop is None:
        init_font()
    return _font_prop


def setup_text_with_font(ax, texts_dict):
    """
    批量设置图表文本的字体

    参数:
        ax: matplotlib axes 对象
        texts_dict: 字典，包含需要设置字本的文本元素
                   {
                       'title': '标题',
                       'xlabel': 'X轴标签',
                       'ylabel': 'Y轴标签',
                       'xtick_labels': ['标签1', '标签2'],
                       'ytick_labels': ['标签1', '标签2'],
                   }
    """
    font_prop = get_font_prop()
    if font_prop is None:
        return

    if 'title' in texts_dict:
        ax.set_title(texts_dict['title'], fontproperties=font_prop)
    if 'xlabel' in texts_dict:
        ax.set_xlabel(texts_dict['xlabel'], fontproperties=font_prop)
    if 'ylabel' in texts_dict:
        ax.set_ylabel(texts_dict['ylabel'], fontproperties=font_prop)

    if 'xtick_labels' in texts_dict:
        ax.set_xticklabels(texts_dict['xtick_labels'], fontproperties=font_prop)
    if 'ytick_labels' in texts_dict:
        ax.set_yticklabels(texts_dict['ytick_labels'], fontproperties=font_prop)

    # 设置图例字体
    legend = ax.get_legend()
    if legend is not None:
        for text in legend.get_texts():
            text.set_fontproperties(font_prop)


# 自动初始化
init_font()


if __name__ == '__main__':
    # 测试
    import numpy as np

    font_prop = get_font_prop()

    fig, ax = plt.subplots(figsize=(8, 6))

    x = np.linspace(0, 10, 100)
    ax.plot(x, np.sin(x), label='正弦波')
    ax.plot(x, np.cos(x), label='余弦波')

    ax.set_title('测试图表', fontproperties=font_prop)
    ax.set_xlabel('横轴', fontproperties=font_prop)
    ax.set_ylabel('纵轴', fontproperties=font_prop)

    legend = ax.get_legend()
    if legend:
        for text in legend.get_texts():
            text.set_fontproperties(font_prop)

    plt.savefig('font_test_final.png', dpi=100, bbox_inches='tight')
    print("✅ 测试图表已保存: font_test_final.png")
