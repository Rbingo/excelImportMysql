#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本，用于将Python应用打包成EXE文件
"""

import os
import sys
import subprocess
import shutil

# 项目根目录
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# 输出目录
OUTPUT_DIR = os.path.join(ROOT_DIR, "dist")

# 清理输出目录
if os.path.exists(OUTPUT_DIR):
    shutil.rmtree(OUTPUT_DIR)
    print(f"已清理输出目录: {OUTPUT_DIR}")

# PyInstaller命令 - 优化版本，保留必要依赖
pyinstaller_cmd = [
    sys.executable,
    "-m", "PyInstaller",
    "--onefile",  # 打包成单个文件
    "--windowed",  # 窗口模式，不显示控制台
    "--name", "ExcelToMySQL",  # 生成的EXE文件名
    "--distpath", OUTPUT_DIR,  # 输出目录
    "--workpath", os.path.join(ROOT_DIR, "build"),  # 工作目录
    "--specpath", ROOT_DIR,  # spec文件目录
    "--add-data", f"{ROOT_DIR}/app;app",  # 包含app目录
    # 移除图标参数，避免权限问题
    # 隐藏导入，确保依赖被正确打包
    "--hidden-import", "pandas",
    "--hidden-import", "pymysql",
    "--hidden-import", "openpyxl",
    "--hidden-import", "requests",
    "--hidden-import", "tkinter.ttk",  # 显式包含ttk，因为我们使用它
    "--strip",  # 移除调试信息，减小文件大小
    # 仅在UPX可用时使用压缩
    # "--upx-dir", ".",  # 注释掉UPX，避免依赖问题
    # 仅排除绝对不需要的模块
    "--exclude-module", "tkinter.tix",  # 排除Tkinter的tix模块（不使用）
    "--exclude-module", "test",  # 排除测试模块
    "--exclude-module", "unittest",  # 排除unittest模块
    "--exclude-module", "doctest",  # 排除doctest模块
    "--exclude-module", "pydoc",  # 排除pydoc模块
    "--exclude-module", "pdb",  # 排除pdb调试模块
    "--exclude-module", "matplotlib",  # 排除matplotlib（不使用）
    "--exclude-module", "scipy",  # 排除scipy（不使用）
    "--exclude-module", "PIL",  # 排除PIL/Pillow（不使用）
    "main.py"  # 主程序文件
]

print(f"执行打包命令: {' '.join(pyinstaller_cmd)}")

# 执行打包命令
try:
    subprocess.run(pyinstaller_cmd, check=True)
    print(f"\n✅ 打包成功！")
    print(f"\n📦 输出文件位置: {OUTPUT_DIR}")
    print(f"\n🎯 可执行文件: {os.path.join(OUTPUT_DIR, 'ExcelToMySQL.exe')}")
    
    # 清理临时文件
    build_dir = os.path.join(ROOT_DIR, "build")
    spec_file = os.path.join(ROOT_DIR, "ExcelToMySQL.spec")
    
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)
        print(f"\n已清理临时构建目录: {build_dir}")
    
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"已清理spec文件: {spec_file}")
        
except subprocess.CalledProcessError as e:
    print(f"\n❌ 打包失败: {e}")
    sys.exit(1)
except Exception as e:
    print(f"\n❌ 打包过程中出现错误: {e}")
    sys.exit(1)
