#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
字体管理工具（统一版）

功能：
1. 诊断字体问题 - 检测matplotlib中文字体配置
2. 修复字体问题 - 自动复制和配置字体
3. 测试字体显示 - 生成测试图表
4. 清除字体缓存 - 清除matplotlib缓存

用法：
    python font_manager.py              # 默认：诊断+修复+测试
    python font_manager.py diagnose     # 仅诊断
    python font_manager.py fix          # 仅修复
    python font_manager.py test         # 仅测试
    python font_manager.py clear-cache # 清除缓存
"""

import sys
import os
import argparse

# 添加项目路径
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, PROJECT_DIR)


def diagnose_font_issue():
    """诊断字体问题"""
    print("\n" + "="*60)
    print("步骤1：诊断字体问题")
    print("="*60)

    try:
        import matplotlib
        import matplotlib.pyplot as plt
        import matplotlib.font_manager as fm

        print(f"✅ matplotlib 版本: {matplotlib.__version__}")
        print(f"✅ 系统平台: {sys.platform}")

        # 检查可用中文字体
        print("\n检查系统中文字体...")
        font_list = fm.findSystemFonts()
        chinese_fonts = [f for f in font_list if any(keyword in f.lower()
                        for keyword in ['hei', 'song', 'kai', 'min', 'pingfang', 'stheit'])]

        if chinese_fonts:
            print(f"✅ 找到 {len(chinese_fonts)} 个中文字体")
            print("  前5个字体:")
            for font in chinese_fonts[:5]:
                print(f"    - {os.path.basename(font)}")
        else:
            print("⚠️  未找到系统中文字体")
            return False

        # 检查项目字体目录
        fonts_dir = os.path.join(PROJECT_DIR, 'price_maintenance_risk_analysis', 'utils', 'fonts')
        if os.path.exists(fonts_dir):
            font_files = [f for f in os.listdir(fonts_dir) if f.endswith(('.ttf', '.ttc'))]
            if font_files:
                print(f"\n✅ 项目字体目录: {fonts_dir}")
                print(f"  包含 {len(font_files)} 个字体文件")
                for font_file in font_files:
                    print(f"    - {font_file}")
            else:
                print(f"\n⚠️  项目字体目录存在但无字体文件: {fonts_dir}")
        else:
            print(f"\n⚠️  项目字体目录不存在: {fonts_dir}")

        # 测试当前配置
        print("\n测试当前字体配置...")
        try:
            from utils.font_manager import get_font_prop
            font_prop = get_font_prop()
            print("✅ 可以获取字体配置")

            # 简单测试
            fig, ax = plt.subplots(figsize=(8, 6))
            ax.text(0.5, 0.5, '中文测试', fontproperties=font_prop, fontsize=20)
            plt.close()
            print("✅ 可以正常使用字体配置")

        except Exception as e:
            print(f"⚠️  字体配置测试失败: {e}")

        return True

    except ImportError as e:
        print(f"❌ matplotlib 未安装: {e}")
        return False
    except Exception as e:
        print(f"❌ 诊断失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def fix_font_issue():
    """修复字体问题"""
    print("\n" + "="*60)
    print("步骤2：修复字体问题")
    print("="*60)

    # 创建字体目录
    fonts_dir = os.path.join(PROJECT_DIR, 'price_maintenance_risk_analysis', 'utils', 'fonts')
    os.makedirs(fonts_dir, exist_ok=True)
    print(f"✅ 字体目录: {fonts_dir}")

    # macOS系统字体
    macos_fonts = [
        '/System/Library/Fonts/STHeiti Medium.ttc',
        '/System/Library/Fonts/PingFang.ttc',
    ]

    print("\n复制系统字体...")
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
        print("\n手动安装方法:")
        print("1. 下载 SimHei.ttf 或其他中文字体")
        print(f"2. 复制到 {fonts_dir}")
        return False

    return True


def test_font_display():
    """测试字体显示"""
    print("\n" + "="*60)
    print("步骤3：测试字体显示")
    print("="*60)

    try:
        import matplotlib.pyplot as plt
        import matplotlib.font_manager as fm

        fonts_dir = os.path.join(PROJECT_DIR, 'price_maintenance_risk_analysis', 'utils', 'fonts')
        font_files = [f for f in os.listdir(fonts_dir) if f.endswith(('.ttf', '.ttc'))]

        print(f"字体目录: {fonts_dir}")
        print(f"找到 {len(font_files)} 个字体文件")

        for font_file in font_files:
            font_path = os.path.join(fonts_dir, font_file)
            print(f"\n测试字体: {font_file}")

            try:
                font_prop = fm.FontProperties(fname=font_path)

                # 生成测试图
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.text(0.5, 0.5, '中文测试 Text 测试',
                       fontsize=30, ha='center', va='center', fontproperties=font_prop)
                ax.set_title('字体测试 Font Test', fontproperties=font_prop, fontsize=16)
                ax.set_xlabel('横轴 X-Axis', fontproperties=font_prop)
                ax.set_ylabel('纵轴 Y-Axis', fontproperties=font_prop)

                # 保存测试图
                test_image = os.path.join(PROJECT_DIR, 'reports', 'images', 'font_test.png')
                os.makedirs(os.path.dirname(test_image), exist_ok=True)
                plt.savefig(test_image, dpi=100, bbox_inches='tight')
                plt.close()

                print(f"✅ 字体测试成功")
                print(f"   测试图: {test_image}")
                print(f"\n   使用方法:")
                print(f"   from utils.font_manager import get_font_prop")
                print(f"   font_prop = get_font_prop()")
                print(f"   ax.set_title('标题', fontproperties=font_prop)")
                return True

            except Exception as e:
                print(f"❌ 测试失败: {e}")

        return False

    except Exception as e:
        print(f"❌ 测试失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def clear_font_cache():
    """清除matplotlib字体缓存"""
    print("\n" + "="*60)
    print("清除matplotlib字体缓存")
    print("="*60)

    try:
        import matplotlib
        import shutil

        # 缓存目录位置
        cache_dir = matplotlib.get_cachedir()
        print(f"缓存目录: {cache_dir}")

        if os.path.exists(cache_dir):
            try:
                shutil.rmtree(cache_dir)
                print("✅ 已清除缓存目录")
                print("   请重启Jupyter Kernel使更改生效")
            except Exception as e:
                print(f"❌ 清除失败: {e}")
        else:
            print("⚠️  缓存目录不存在")

        # 清除fontManager的缓存
        try:
            import matplotlib.font_manager
            matplotlib.font_manager._rebuild()
            print("✅ 已重建字体管理器缓存")
        except Exception as e:
            print(f"⚠️  重建缓存失败: {e}")

        return True

    except ImportError:
        print("❌ matplotlib 未安装")
        return False
    except Exception as e:
        print(f"❌ 清除缓存失败: {e}")
        return False


def main_full_workflow():
    """完整工作流：诊断+修复+测试"""
    print("""
╔══════════════════════════════════════════════════════════╗
║              中文字体管理工具                                ║
║            解决matplotlib中文显示问题                     ║
╚══════════════════════════════════════════════════════════╝
""")

    print(f"系统: {sys.platform}")
    print(f"Python: {sys.version.split()[0]}")

    # 步骤1: 诊断
    diagnose_result = diagnose_font_issue()

    if not diagnose_result:
        print("\n❌ 诊断失败，请检查matplotlib安装")
        return False

    # 步骤2: 修复
    fix_result = fix_font_issue()
    if not fix_result:
        print("\n❌ 修复失败，请手动安装字体")
        return False

    # 步骤3: 测试
    test_result = test_font_display()
    if not test_result:
        print("\n❌ 测试失败，请检查字体文件")
        return False

    # 步骤4: 清除缓存
    print("\n建议清除matplotlib缓存以生效...")
    print("运行: python font_manager.py clear-cache")

    print("\n" + "="*60)
    print("✅ 字体配置完成！")
    print("="*60)
    print("\n在Jupyter Notebook中使用:")
    print("  from utils.font_manager import get_font_prop")
    print("  font_prop = get_font_prop()")
    print("  ax.set_title('标题', fontproperties=font_prop)")

    return True


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='中文字体管理工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
使用示例:
  python font_manager.py              # 默认：完整流程（诊断+修复+测试）
  python font_manager.py diagnose     # 仅诊断
  python font_manager.py fix          # 仅修复
  python font_manager.py test         # 仅测试
  python font_manager.py clear-cache # 清除缓存
        '''
    )

    parser.add_argument(
        'command',
        nargs='?',
        choices=['diagnose', 'fix', 'test', 'clear-cache'],
        help='操作命令（默认：完整流程）'
    )

    args = parser.parse_args()

    if args.command is None:
        # 默认运行完整流程
        main_full_workflow()
    elif args.command == 'diagnose':
        diagnose_font_issue()
    elif args.command == 'fix':
        fix_font_issue()
    elif args.command == 'test':
        test_font_display()
    elif args.command == 'clear-cache':
        clear_font_cache()


if __name__ == '__main__':
    main()
