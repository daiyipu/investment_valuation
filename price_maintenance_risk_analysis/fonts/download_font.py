#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
中文字体下载脚本
下载开源中文字体到本地
"""

import os
import urllib.request
import sys


def download_font(url, dest_path):
    """下载字体文件"""
    print(f"正在下载: {url}")
    print(f"保存到: {dest_path}")

    try:
        def progress(block_num, block_size, total_size):
            downloaded = block_num * block_size
            percent = int((downloaded / total_size) * 100) if total_size > 0 else 0
            sys.stdout.write(f"\r进度: {percent}% ({downloaded}/{total_size} bytes)")
            sys.stdout.flush()

        urllib.request.urlretrieve(url, dest_path, progress)
        print(f"\n✅ 下载完成: {dest_path}")
        return True
    except Exception as e:
        print(f"\n❌ 下载失败: {e}")
        return False


def main():
    print("="*60)
    print("中文字体下载工具")
    print("="*60)

    # 确保目录存在
    font_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(font_dir, exist_ok=True)

    # 使用开源字体（来自 Google Noto Sans CJK 或其他开源项目）
    # 这里提供一个更小的、使用友好的方案

    print("\n请选择字体下载方式:")
    print("1. 从网络下载（需要联网）")
    print("2. 本地复制（从系统复制）")
    print("3. 显示手动下载说明")

    choice = input("\n请输入选项 (1/2/3): ").strip()

    if choice == '1':
        print("\n⚠️ 由于字体文件较大且版权问题，")
        print("建议使用手动下载或系统复制方式。")
        print("\n推荐的开源中文字体:")
        print("- Noto Sans CJK: https://github.com/googlefonts/noto-cjk")
        print("- 文泉驿字体: https://github.com/adobe-fonts/source-han-sans")
        print("\n或者直接复制系统字体（选择选项2）")

    elif choice == '2':
        print("\n=== 复制系统字体 ===")
        system = sys.platform

        if system == 'darwin':
            print("\nmacOS 系统字体:")
            src_fonts = [
                '/System/Library/Fonts/STHeiti Medium.ttc',
                '/System/Library/Fonts/PingFang.ttc'
            ]
            for src in src_fonts:
                if os.path.exists(src):
                    dest = os.path.join(font_dir, os.path.basename(src))
                    try:
                        import shutil
                        shutil.copy(src, dest)
                        print(f"✅ 已复制: {os.path.basename(src)}")
                    except Exception as e:
                        print(f"❌ 复制失败: {e}")
                    break

        elif system.startswith('linux'):
            print("\nLinux 系统字体:")
            src_fonts = [
                '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
                '/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc'
            ]
            for src in src_fonts:
                if os.path.exists(src):
                    dest = os.path.join(font_dir, os.path.basename(src))
                    try:
                        import shutil
                        shutil.copy(src, dest)
                        print(f"✅ 已复制: {os.path.basename(src)}")
                    except Exception as e:
                        print(f"❌ 复制失败: {e}")
                    break

            print("\n如果找不到字体，请先安装:")
            print("  sudo apt-get install fonts-wqy-microhei fonts-noto-cjk")

        elif system == 'win32':
            print("\nWindows 系统字体:")
            src_fonts = [
                'C:\\Windows\\Fonts\\simhei.ttf',
                'C:\\Windows\\Fonts\\simsun.ttc'
            ]
            for src in src_fonts:
                if os.path.exists(src):
                    dest = os.path.join(font_dir, os.path.basename(src))
                    try:
                        import shutil
                        shutil.copy(src, dest)
                        print(f"✅ 已复制: {os.path.basename(src)}")
                    except Exception as e:
                        print(f"❌ 复制失败: {e}")
                    break

    elif choice == '3':
        print("\n=== 手动下载说明 ===")
        print("\n方法 1 - 从 GitHub 下载:")
        print("  https://github.com/StellarCN/scp_zh/blob/master/fonts/SimHei.ttf")
        print("  下载后放到此目录")

        print("\n方法 2 - Windows 系统自带:")
        print("  从 C:\\Windows\\Fonts\\ 复制 simhei.ttf")

        print("\n方法 3 - macOS 系统自带:")
        print("  从 /System/Library/Fonts/ 复制 STHeiti*.ttc")

        print("\n方法 4 - Linux 安装:")
        print("  sudo apt-get install fonts-wqy-microhei")

    print("\n完成后，请重启 Jupyter Notebook")

    # 检查当前目录的字体文件
    print("\n=== 当前字体目录 ===")
    font_files = [f for f in os.listdir(font_dir) if f.endswith(('.ttf', '.ttc', '.otf'))]
    if font_files:
        print(f"找到 {len(font_files)} 个字体文件:")
        for f in font_files:
            print(f"  - {f}")
        print("\n✅ 字体已就绪，可以使用！")
    else:
        print("暂无字体文件")


if __name__ == '__main__':
    main()
