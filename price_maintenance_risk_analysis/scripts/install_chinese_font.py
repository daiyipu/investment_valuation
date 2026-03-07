#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
中文字体自动安装脚本
适用于 Linux/vnpy 环境
"""

import os
import sys
import urllib.request
import zipfile
import platform


def install_font_linux():
    """Linux 环境下安装中文字体"""
    print("检测到 Linux 系统，开始配置中文字体...")

    # 检查字体是否已安装
    try:
        import matplotlib.font_manager as fm
        available = [f.name for f in fm.fontManager.ttflist]
        chinese_fonts = [f for f in available if any(kw in f.lower() for kw in ['noto', 'wenquanyi', 'wqy', 'cjk'])]

        if chinese_fonts:
            print(f"✅ 已有中文字体: {chinese_fonts[:3]}")
            return True
    except Exception as e:
        print(f"检查字体时出错: {e}")

    print("\n请选择安装方式:")
    print("1. 使用系统包管理器安装（推荐）")
    print("2. 手动下载字体文件")

    choice = input("\n请输入选项 (1/2): ").strip()

    if choice == '1':
        print("\n请运行以下命令安装字体:")
        print("\nUbuntu/Debian:")
        print("  sudo apt-get update")
        print("  sudo apt-get install -y fonts-noto-cjk fonts-wqy-microhei")
        print("\nCentOS/RedHat:")
        print("  sudo yum install -y wqy-microhei-fonts wqy-zenhei-fonts")
        print("\nArch Linux:")
        print("  sudo pacman -S noto-fonts-cjk")
        print("\n安装完成后，请运行: python diagnose_font.py clear-cache")

    elif choice == '2':
        print("\n手动安装步骤:")
        print("1. 下载字体文件（例如 Source Han Sans）")
        print("   下载地址: https://github.com/adobe-fonts/source-han-sans/releases")
        print("2. 创建字体目录:")
        print(f"   mkdir -p {os.path.expanduser('~/.fonts')}")
        print("3. 将字体文件复制到 ~/.fonts/ 目录")
        print("4. 更新字体缓存:")
        print("   fc-cache -fv")
        print("5. 清除 matplotlib 缓存:")
        print("   python diagnose_font.py clear-cache")

    return False


def install_font_macos():
    """macOS 系统字体配置"""
    print("检测到 macOS 系统")

    print("\nmacOS 系统通常已内置中文字体。")
    print("如果仍显示乱码，请尝试:")

    print("\n1. 检查字体是否启用:")
    print("   打开'字体册'应用（应用程序 -> 字体册）")
    print("   搜索'PingFang'或'华文黑体'")

    print("\n2. 如果未启用，双击字体文件并点击'安装'")

    print("\n3. 清除 matplotlib 缓存:")
    print("   python diagnose_font.py clear-cache")

    return False


def install_font_windows():
    """Windows 系统字体配置"""
    print("检测到 Windows 系统")

    print("\nWindows 系统通常已安装中文字体。")
    print("如果仍显示乱码，请尝试:")

    print("\n1. 确认系统字体是否正常:")
    print("   C:\\Windows\\Fonts\\ 目录下应有 'msyh.ttc' (微软雅黑)")

    print("\n2. 如果没有，请下载安装:")
    print("   从其他 Windows 电脑复制 msyh.ttc 到 C:\\Windows\\Fonts\\")

    print("\n3. 清除 matplotlib 缓存:")
    print("   python diagnose_font.py clear-cache")

    return False


def main():
    print("""
╔══════════════════════════════════════════════════════════╗
║             中文字体安装助手                              ║
╚══════════════════════════════════════════════════════════╝
""")

    system = platform.system()

    if system == 'Linux':
        install_font_linux()
    elif system == 'Darwin':
        install_font_macos()
    elif system == 'Windows':
        install_font_windows()
    else:
        print(f"未知系统: {system}")
        print("请手动安装中文字体")


if __name__ == '__main__':
    main()
