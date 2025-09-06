#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
构建可执行文件的脚本
使用PyInstaller将Excel对比工具打包成exe文件
"""

import os
import sys
import subprocess

def install_pyinstaller():
    """安装PyInstaller"""
    try:
        import PyInstaller
        print("PyInstaller已安装")
        return True
    except ImportError:
        print("正在安装PyInstaller...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("PyInstaller安装成功")
            return True
        except subprocess.CalledProcessError:
            print("PyInstaller安装失败")
            return False

def build_exe():
    """构建exe文件"""
    if not install_pyinstaller():
        return False
    
    # 获取当前目录
    current_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(current_dir, "excel_compare_tool.py")
    
    if not os.path.exists(script_path):
        print(f"错误：找不到文件 {script_path}")
        return False
    
    # PyInstaller命令参数
    cmd = [
        "pyinstaller",
        "--onefile",  # 打包成单个exe文件
        "--windowed",  # 不显示控制台窗口
        "--name=Excel数据对比工具",  # 设置exe文件名
        "--icon=NONE",  # 可以后续添加图标
        "--add-data=requirements.txt;.",  # 包含requirements.txt
        script_path
    ]
    
    try:
        print("开始构建exe文件...")
        print(f"执行命令: {' '.join(cmd)}")
        
        # 执行PyInstaller命令
        result = subprocess.run(cmd, cwd=current_dir, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("\n构建成功！")
            exe_path = os.path.join(current_dir, "dist", "Excel数据对比工具.exe")
            if os.path.exists(exe_path):
                print(f"可执行文件位置: {exe_path}")
                print("\n使用说明:")
                print("1. 双击 'Excel数据对比工具.exe' 即可启动程序")
                print("2. 无需安装Python环境")
                print("3. 可以将exe文件复制到任何Windows电脑上运行")
            else:
                print("警告：exe文件可能未正确生成")
        else:
            print("\n构建失败！")
            print("错误信息:")
            print(result.stderr)
            
    except Exception as e:
        print(f"构建过程中出现错误: {e}")
        return False
    
    return result.returncode == 0

def clean_build_files():
    """清理构建过程中产生的临时文件"""
    import shutil
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 要清理的目录和文件
    cleanup_items = [
        os.path.join(current_dir, "build"),
        os.path.join(current_dir, "Excel数据对比工具.spec")
    ]
    
    for item in cleanup_items:
        if os.path.exists(item):
            try:
                if os.path.isdir(item):
                    shutil.rmtree(item)
                    print(f"已删除目录: {item}")
                else:
                    os.remove(item)
                    print(f"已删除文件: {item}")
            except Exception as e:
                print(f"清理 {item} 时出错: {e}")

if __name__ == "__main__":
    print("Excel数据对比工具 - 可执行文件构建器")
    print("=" * 50)
    
    # 构建exe
    success = build_exe()
    
    if success:
        print("\n是否清理构建临时文件？(y/n): ", end="")
        choice = input().lower().strip()
        if choice in ['y', 'yes', '是']:
            clean_build_files()
            print("临时文件清理完成")
    
    print("\n按任意键退出...")
    input()