# merge_pdf_gui.py
# -*- coding: utf-8 -*-

"""
带简单 GUI 的 PDF 合并工具（Windows / macOS / Linux）。
依赖: PyPDF2（pip install PyPDF2）
保存为 merge_pdf_gui.py 后运行：python merge_pdf_gui.py
"""

import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfMerger, PdfReader
import time
import sys
import os

# ---------------- 默认参数（可在 GUI 中修改） ----------------
DEFAULT_RECURSIVE = False
DEFAULT_SORT_BY = "name"  # "name" or "mtime"
DEFAULT_SKIP_ENCRYPTED = True
# ------------------------------------------------------------

class PDFMergerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 合并小工具")
        self.geometry("720x420")
        self.resizable(False, False)

        self.input_dir = tk.StringVar()
        self.output_file = tk.StringVar()
        self.recursive = tk.BooleanVar(value=DEFAULT_RECURSIVE)
        self.sort_by = tk.StringVar(value=DEFAULT_SORT_BY)
        self.skip_encrypted = tk.BooleanVar(value=DEFAULT_SKIP_ENCRYPTED)
        self.file_list = []

        self._build_ui()

    def _build_ui(self):
        pad = 8

        frame_top = ttk.Frame(self, padding=pad)
        frame_top.pack(fill="x")

        # 输入文件夹
        ttk.Label(frame_top, text="输入文件夹:").grid(column=0, row=0, sticky="w")
        ttk.Entry(frame_top, textvariable=self.input_dir, width=60).grid(column=1, row=0, padx=6)
        ttk.Button(frame_top, text="选择", command=self.choose_input_dir).grid(column=2, row=0)

        # 输出文件
        ttk.Label(frame_top, text="输出文件:").grid(column=0, row=1, sticky="w", pady=(6,0))
        ttk.Entry(frame_top, textvariable=self.output_file, width=60).grid(column=1, row=1, padx=6, pady=(6,0))
        ttk.Button(frame_top, text="保存为...", command=self.choose_output_file).grid(column=2, row=1, pady=(6,0))

        # 选项
        frame_opts = ttk.Frame(self, padding=pad)
        frame_opts.pack(fill="x")
        ttk.Checkbutton(frame_opts, text="递归子目录", variable=self.recursive).grid(column=0, row=0, sticky="w")
        ttk.Checkbutton(frame_opts, text="跳过加密文件", variable=self.skip_encrypted).grid(column=1, row=0, sticky="w", padx=10)
        ttk.Label(frame_opts, text="排序方式:").grid(column=2, row=0, sticky="e")
        sort_combo = ttk.Combobox(frame_opts, values=["name", "mtime"], width=8, state="readonly", textvariable=self.sort_by)
        sort_combo.grid(column=3, row=0, padx=(6,0))
        sort_combo.set(self.sort_by.get())

        # 中间：文件列表与操作按钮
        frame_mid = ttk.Frame(self, padding=pad)
        frame_mid.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(frame_mid, height=12, width=90)
        self.listbox.pack(side="left", padx=(0,6), pady=(6,0))
        scrollbar = ttk.Scrollbar(frame_mid, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="left", fill="y", pady=(6,0))
        self.listbox.config(yscrollcommand=scrollbar.set)

        frame_buttons = ttk.Frame(frame_mid)
        frame_buttons.pack(side="left", fill="y", padx=(6,0))
        ttk.Button(frame_buttons, text="刷新文件列表", command=self.refresh_file_list).pack(fill="x", pady=3)
        ttk.Button(frame_buttons, text="清空列表", command=self.clear_list).pack(fill="x", pady=3)
        ttk.Button(frame_buttons, text="开始合并", command=self.start_merge_thread).pack(fill="x", pady=20)
        ttk.Button(frame_buttons, text="退出", command=self.quit).pack(fill="x", pady=3)

        # 进度条与状态栏
        frame_bot = ttk.Frame(self, padding=pad)
        frame_bot.pack(fill="x")
        self.progress = ttk.Progressbar(frame_bot, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=6, pady=(0,6))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(frame_bot, textvariable=self.status_var).pack(anchor="w", padx=6)

    def choose_input_dir(self):
        d = filedialog.askdirectory(title="选择包含 PDF 的文件夹")
        if d:
            self.input_dir.set(d)
            self.refresh_file_list()

    def choose_output_file(self):
        f = filedialog.asksaveasfilename(title="保存合并的 PDF 为", defaultextension=".pdf", filetypes=[("PDF 文件","*.pdf")])
        if f:
            self.output_file.set(f)

    def collect_pdfs(self, directory: Path, recursive: bool):
        pattern = "**/*.pdf" if recursive else "*.pdf"
        return [p for p in directory.glob(pattern) if p.is_file()]

    def refresh_file_list(self):
        self.listbox.delete(0, tk.END)
        self.file_list = []
        dir_path = self.input_dir.get().strip()
        if not dir_path:
            self.status_var.set("请先选择输入文件夹。")
            return
        p = Path(dir_path)
        if not p.exists() or not p.is_dir():
            self.status_var.set("输入路径无效。")
            return
        pdfs = self.collect_pdfs(p, self.recursive.get())
        if not pdfs:
            self.status_var.set("未找到 PDF 文件。")
            return

        # 排序
        if self.sort_by.get() == "name":
            pdfs.sort(key=lambda x: x.name.lower())
        else:
            pdfs.sort(key=lambda x: x.stat().st_mtime)

        for f in pdfs:
            self.listbox.insert(tk.END, f.name)
        self.file_list = pdfs
        self.status_var.set(f"找到 {len(self.file_list)} 个 PDF。")

    def clear_list(self):
        self.listbox.delete(0, tk.END)
        self.file_list = []
        self.status_var.set("列表已清空。")

    def start_merge_thread(self):
        # 不要在主线程做合并，避免 GUI 假死
        t = threading.Thread(target=self.merge_pdfs, daemon=True)
        t.start()

    def merge_pdfs(self):
        try:
            # 校验
            if not self.file_list:
                # 如果列表为空尝试自动刷新
                self.refresh_file_list()
                if not self.file_list:
                    messagebox.showwarning("没有 PDF", "未找到待合并的 PDF 文件。")
                    return

            out_path = self.output_file.get().strip()
            if not out_path:
                messagebox.showwarning("输出路径", "请先选择输出文件路径。")
                return
            out_path = Path(out_path)
            out_path.parent.mkdir(parents=True, exist_ok=True)

            self.progress["maximum"] = len(self.file_list)
            self.progress["value"] = 0
            self.status_var.set("开始合并...")
            merger = PdfMerger()
            appended = 0

            for idx, pdf_path in enumerate(self.file_list, start=1):
                self.status_var.set(f"处理中：{pdf_path.name} （{idx}/{len(self.file_list)}）")
                self.update_idletasks()
                try:
                    reader = PdfReader(str(pdf_path))
                    if getattr(reader, "is_encrypted", False):
                        if self.skip_encrypted.get():
                            self.listbox.itemconfig(idx-1, fg="gray")
                            self.append_log(f"跳过加密：{pdf_path.name}")
                            self.progress["value"] = idx
                            continue
                        else:
                            try:
                                reader.decrypt("")  # 尝试空密码
                            except Exception:
                                self.append_log(f"无法解密，跳过：{pdf_path.name}")
                                self.progress["value"] = idx
                                continue

                    merger.append(str(pdf_path))
                    appended += 1
                    self.append_log(f"已加入：{pdf_path.name}")
                except Exception as e:
                    self.append_log(f"加入失败（跳过）：{pdf_path.name}，原因：{e}")
                finally:
                    self.progress["value"] = idx
                    self.update_idletasks()
                    # small sleep 仅用于 UI 更流畅（可移除）
                    time.sleep(0.01)

            if appended == 0:
                messagebox.showerror("失败", "没有成功加入任何 PDF，未生成输出文件。")
                self.status_var.set("合并失败：没有加入任何文件。")
            else:
                merger.write(str(out_path))
                merger.close()
                messagebox.showinfo("完成", f"合并完成，共加入 {appended} 个 PDF。\n输出文件：{out_path}")
                self.status_var.set(f"合并完成：{out_path}")
        except Exception as ex:
            traceback.print_exc()
            messagebox.showerror("错误", f"合并过程中发生错误：\n{ex}")
            self.status_var.set("合并出错。")

    def append_log(self, text):
        # 将 log 显示在 listbox 的下方（直接在 listbox 中标注）
        # 这里不新增 UI 元素，仅在控制台打印，需更多可把日志放到 Text 控件里
        print(text)

def main():
    app = PDFMergerGUI()
    # 如果启动时需要设置默认目录/输出，可在这里赋值，例如：
    # app.input_dir.set(r"C:\pdfs")
    # app.output_file.set(r"C:\out\merged.pdf")
    app.mainloop()

if __name__ == "__main__":
    main()
