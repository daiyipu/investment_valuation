# -*- coding: utf-8 -*-
"""
超简单字体修复 - 直接在 notebook 中使用
复制这段代码到 notebook 第一个单元格运行
"""

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import os

# === 字体配置 ===
# macOS 系统字体路径（直接指定，不依赖检测）
FONT_PATH = '/System/Library/Fonts/STHeiti Medium.ttc'

# 备用字体路径
if not os.path.exists(FONT_PATH):
    FONT_PATH = '/System/Library/Fonts/PingFang.ttc'

# 创建字体属性
try:
    font_prop = fm.FontProperties(fname=FONT_PATH)
    print(f"✅ 字体加载成功: {os.path.basename(FONT_PATH)}")
except Exception as e:
    print(f"❌ 字体加载失败: {e}")
    font_prop = None
    print("⚠️ 图表中文可能显示为方框")

# 设置全局字体（可选，可能有副作用）
# plt.rcParams['font.sans-serif'] = [font_prop.get_name()]
plt.rcParams['axes.unicode_minus'] = False

# === 使用说明 ===
print("\n使用方法:")
print("在绘图时，为每个中文文本添加 fontproperties=font_prop")
print("\n示例:")
print("  ax.set_title('标题', fontproperties=font_prop)")
print("  ax.set_xlabel('X轴', fontproperties=font_prop)")
print("  ax.text(x, y, '文本', fontproperties=font_prop)")

# === 测试代码 ===
def test_font():
    """测试字体是否正常"""
    if font_prop is None:
        print("⚠️ 字体未加载，跳过测试")
        return

    import numpy as np

    fig, ax = plt.subplots(figsize=(10, 6))

    x = np.linspace(0, 10, 100)
    ax.plot(x, np.sin(x), label='正弦波')
    ax.plot(x, np.cos(x), label='余弦波')

    # 注意：每个中文文本都需要添加 fontproperties=font_prop
    ax.set_title('字体测试 - 如果看到中文说明成功', fontproperties=font_prop, fontsize=16)
    ax.set_xlabel('横轴', fontproperties=font_prop)
    ax.set_ylabel('纵轴', fontproperties=font_prop)

    # 设置图例字体
    legend = ax.get_legend()
    if legend:
        for text in legend.get_texts():
            text.set_fontproperties(font_prop)

    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig('font_test_direct.png', dpi=100, bbox_inches='tight')
    plt.show()

    print("✅ 测试图片已保存: font_test_direct.png")

# 自动运行测试
if __name__ != '__main__':
    test_font()
