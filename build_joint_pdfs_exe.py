# -*- coding: utf-8 -*-
"""
打包脚本：将 joint_pdfs_exe.py 打包为 exe 可执行文件
使用方法：python build_joint_pdfs_exe.py
"""

import PyInstaller.__main__
import os

# 获取当前脚本所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# PyInstaller 参数
args = [
    'joint_pdfs_exe.py',  # 主程序文件
    '--name=PDF和图片合并工具',  # 生成的exe名称
    '--onefile',  # 打包为单个exe文件
    '--windowed',  # 不显示控制台窗口（GUI程序）
    '--icon=NONE',  # 如果有图标文件，可以指定：--icon=icon.ico
    # 确保包含所有必要的模块
    '--hidden-import=tkinter',
    '--hidden-import=tkinter.ttk',
    '--hidden-import=tkinter.filedialog',
    '--hidden-import=tkinter.messagebox',
    '--hidden-import=PyPDF2',
    '--hidden-import=fitz',  # PyMuPDF
    '--clean',  # 清理临时文件
]

# 执行打包
PyInstaller.__main__.run(args)

print("\n打包完成！")
print(f"exe文件位置：{os.path.join(current_dir, 'dist', 'PDF和图片合并工具.exe')}")
