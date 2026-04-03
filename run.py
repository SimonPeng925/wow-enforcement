"""
WOW English 维权信息提取 - 一键启动器
使用方法：
  双击运行此文件，或在命令行运行 python run.bat
"""

import subprocess
import sys
import os

def main():
    print("=" * 60)
    print("WOW English 电商维权信息提取工具")
    print("=" * 60)

    # 检查依赖
    print("\n[1/3] 检查依赖...")
    try:
        import playwright
        import openpyxl
        print("    OK - 依赖已安装")
    except ImportError as e:
        print(f"    需要安装依赖: {e}")
        print("    运行: pip install playwright openpyxl")
        input("\n按回车退出...")
        return

    # 获取链接
    print("\n[2/3] 请输入商品链接（直接回车从剪贴板读取）:")
    url = input("    >>> ").strip()

    if not url:
        try:
            import pyperclip
            url = pyperclip.paste()
            if url:
                print(f"    已从剪贴板读取: {url[:50]}...")
        except ImportError:
            print("    未检测到链接，请手动输入")
            url = input("    >>> ").strip()

    if not url:
        print("\n未提供链接，退出。")
        input("\n按回车退出...")
        return

    # 运行提取
    print("\n[3/3] 启动提取程序...")
    print("-" * 60)

    cmd = [sys.executable, "extract_product.py", url]
    result = subprocess.run(cmd, cwd=os.path.dirname(os.path.abspath(__file__)))

    print("-" * 60)
    print("\n提取完成！")
    input("\n按回车退出...")

if __name__ == "__main__":
    main()
