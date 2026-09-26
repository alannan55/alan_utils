# -*- coding: utf-8 -*-
"""
打包脚本：将 compare_excel.py 打包为 exe 可执行文件
使用方法：python build_compare_excel_exe.py
"""

import PyInstaller.__main__
import os

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# PyInstaller 参数
args = [
    'compare_excel.py',  # 主程序文件
    '--name=比较Excel表格并标记差异',  # 生成的exe名称
    '--onefile',  # 打包为单个exe文件
    '--windowed',  # 不显示控制台窗口（GUI程序）
    '--icon=NONE',  # 如果有图标文件，可以指定：--icon=icon.ico
    '--hidden-import=openpyxl',  # 确保包含openpyxl
    '--hidden-import=tkinter',  # 确保包含tkinter
    '--hidden-import=tkinter.ttk',  # 确保包含ttk
    '--hidden-import=tkinter.filedialog',  # 确保包含filedialog
    '--hidden-import=tkinter.messagebox',  # 确保包含messagebox
    '--clean',  # 清理临时文件
]

# 执行打包
PyInstaller.__main__.run(args)

print("\n打包完成！")
print(f"exe文件位置：{os.path.join(current_dir, 'dist', '比较Excel表格并标记差异.exe')}")
