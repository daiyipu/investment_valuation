#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
字体诊断脚本
用于检测和修复中文字体显示问题
"""

import sys
import os

# 添加项目路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils.font_config import diagnose_font_issue, clear_font_cache, setup_chinese_font


def main():
    print("""
╔══════════════════════════════════════════════════════════╗
║             中文字体诊断工具                               ║
║            用于解决 Jupyter 中文乱码问题                    ║
╚══════════════════════════════════════════════════════════╝
""")

    if len(sys.argv) > 1:
        command = sys.argv[1]
        if command == 'clear-cache':
            print("清除 matplotlib 字体缓存...")
            clear_font_cache()
            return
        elif command == 'test':
            print("运行字体测试...")
            from utils.font_config import test_chinese_font
            test_chinese_font()
            return

    # 默认运行完整诊断
    result = diagnose_font_issue()

    print("\n" + "="*60)
    print("诊断完成！")
    print("="*60)

    if not result['available_fonts']:
        print("\n⚠️ 未检测到中文字体，请按照上述说明安装字体")
        print("\n安装后，请运行以下命令清除缓存:")
        print("  python diagnose_font.py clear-cache")
    else:
        print("\n✅ 已配置中文字体，可以在 Jupyter Notebook 中使用了")


if __name__ == '__main__':
    main()
