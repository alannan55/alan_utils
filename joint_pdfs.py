#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
合并指定目录下所有 PDF 为一个 PDF 文件。
修改下方变量后直接运行。
依赖: PyPDF2 (pip install PyPDF2)
"""

from pathlib import Path
from PyPDF2 import PdfMerger, PdfReader
import sys

# ----------------- 参数（按需修改） -----------------
input_dir = Path(r"C:\Users\Alan\Downloads\发票批量下载_20251128175447")   # 要扫描的目录（字符串或 Path），例如 r"C:\pdfs" 或 "/home/user/pdfs"
output_pdf = Path(r"C:\Users\Alan\Downloads\发票批量下载_20251128175447\1128发票合集.pdf")  # 输出合并后的 PDF 路径
recursive = False    # 是否递归子目录（True / False）
sort_by = "name"     # 排序方式: "name"（按文件名） 或 "mtime"（按修改时间）
skip_encrypted = True  # 是否跳过加密的 PDF（True 跳过，False 尝试合并但可能失败）
# ----------------------------------------------------

def collect_pdfs(directory: Path, recursive: bool):
    pattern = "**/*.pdf" if recursive else "*.pdf"
    return [p for p in directory.glob(pattern) if p.is_file()]

def main():
    # 检查输入目录
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"错误：输入目录不存在或不是目录：{input_dir}")
        sys.exit(1)

    pdf_list = collect_pdfs(input_dir, recursive)
    if not pdf_list:
        print("未找到任何 PDF 文件。")
        sys.exit(0)

    # 排序
    if sort_by == "name":
        pdf_list.sort(key=lambda p: p.name.lower())
    elif sort_by == "mtime":
        pdf_list.sort(key=lambda p: p.stat().st_mtime)
    else:
        print(f"未知的 sort_by 值：{sort_by}，使用文件名排序。")
        pdf_list.sort(key=lambda p: p.name.lower())

    print(f"找到 {len(pdf_list)} 个 PDF，开始合并...")

    merger = PdfMerger()
    appended = 0
    try:
        for pdf_path in pdf_list:
            try:
                # 检查是否加密
                reader = PdfReader(str(pdf_path))
                if getattr(reader, "is_encrypted", False):
                    if skip_encrypted:
                        print(f"跳过加密文件：{pdf_path.name}")
                        continue
                    else:
                        # 尝试解密空密码（有时可行）
                        try:
                            reader.decrypt("")  # 若有密码需要手动提供
                        except Exception:
                            print(f"无法读取加密文件：{pdf_path.name} （跳过）")
                            continue

                merger.append(str(pdf_path))
                appended += 1
                print(f"已加入：{pdf_path.name}")
            except Exception as e:
                print(f"加入失败（跳过）：{pdf_path.name}，原因：{e}")
                continue

        if appended == 0:
            print("没有成功加入任何 PDF，未生成输出文件。")
            sys.exit(1)

        # 确保父目录存在
        output_pdf.parent.mkdir(parents=True, exist_ok=True)
        merger.write(str(output_pdf))
        print(f"合并完成，共加入 {appended} 个 PDF。输出文件：{output_pdf}")
    finally:
        try:
            merger.close()
        except Exception:
            pass

if __name__ == "__main__":
    main()
