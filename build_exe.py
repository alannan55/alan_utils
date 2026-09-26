# -*- coding: utf-8 -*-
"""
打包脚本：将 split_supplier_xlsx.py 打包为 exe 可执行文件
使用方法：python build_exe.py
"""

import PyInstaller.__main__
import os
import shutil

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# PyInstaller 参数
args = [
    'split_supplier_xlsx.py',  # 主程序文件
    '--name=按供应商拆分Excel表格',  # 生成的exe名称
    '--onefile',  # 打包为单个exe文件
    '--windowed',  # 不显示控制台窗口（GUI程序）
    '--icon=NONE',  # 如果有图标文件，可以指定：--icon=icon.ico
    '--add-data=supplier_shortnames.json;.',  # 包含JSON配置文件（Windows使用分号分隔）
    '--hidden-import=openpyxl',  # 确保包含openpyxl
    '--hidden-import=tkinter',  # 确保包含tkinter
    '--clean',  # 清理临时文件
]

# 执行打包
PyInstaller.__main__.run(args)

dist_config_file = os.path.join(current_dir, 'dist', 'supplier_shortnames.json')
source_config_file = os.path.join(current_dir, 'supplier_shortnames.json')
if os.path.exists(source_config_file):
    shutil.copy2(source_config_file, dist_config_file)

print("\n打包完成！")
print(f"exe文件位置：{os.path.join(current_dir, 'dist', '按供应商拆分Excel表格.exe')}")
print(f"外部配置文件位置：{dist_config_file}")
