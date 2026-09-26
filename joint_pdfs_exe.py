# merge_pdf_gui.py
# -*- coding: utf-8 -*-

"""
带简单 GUI 的 PDF 和图片合并工具（Windows / macOS / Linux）。
支持合并 PDF 文件和常见图片格式（jpg, jpeg, png, bmp, tif, tiff）。
所有文件在合并时会自动调整到统一的标准尺寸（A4）。
依赖: PyPDF2（pip install PyPDF2）
      Pillow（pip install Pillow）- 用于图片转PDF和调整大小
      PyMuPDF（pip install PyMuPDF）- 用于 PDF 转图片和调整大小
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
import subprocess
import platform
from io import BytesIO

# 尝试导入 Pillow，如果未安装则提示
try:
    from PIL import Image
    PIL_AVAILABLE = True
    # 兼容性：检查是否有 Resampling 属性（Pillow 10.0+）
    try:
        RESAMPLE_MODE = Image.Resampling.LANCZOS
    except AttributeError:
        RESAMPLE_MODE = Image.LANCZOS  # 旧版本使用
except ImportError:
    PIL_AVAILABLE = False
    RESAMPLE_MODE = None

# 尝试导入 PyMuPDF (fitz)，用于 PDF 转图片
try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

# ---------------- 默认参数（可在 GUI 中修改） ----------------
DEFAULT_RECURSIVE = False
DEFAULT_SORT_BY = "name"  # "name" or "mtime"
DEFAULT_SKIP_ENCRYPTED = True
# 标准页面尺寸（A4，单位：像素，300 DPI）
STANDARD_PAGE_SIZE = (2480, 3508)  # A4 尺寸 (210mm x 297mm) @ 300 DPI
# ------------------------------------------------------------

class PDFMergerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 和图片合并小工具")
        self.geometry("800x420")
        self.resizable(False, False)

        self.input_dir = tk.StringVar()
        self.output_file = tk.StringVar()
        self.recursive = tk.BooleanVar(value=DEFAULT_RECURSIVE)
        self.sort_by = tk.StringVar(value=DEFAULT_SORT_BY)
        self.skip_encrypted = tk.BooleanVar(value=DEFAULT_SKIP_ENCRYPTED)
        self.file_list = []
        self.last_merged_pdf = None  # 保存最后合并的PDF路径
        self.merge_thread = None  # 保存合并线程引用

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)  # 处理窗口关闭事件

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
        self.listbox = tk.Listbox(frame_mid, height=12, width=85)
        self.listbox.pack(side="left", padx=(0,6), pady=(6,0))
        scrollbar = ttk.Scrollbar(frame_mid, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="left", fill="y", pady=(6,0))
        self.listbox.config(yscrollcommand=scrollbar.set)

        frame_buttons = ttk.Frame(frame_mid)
        frame_buttons.pack(side="left", fill="y", padx=(6,0))
        ttk.Button(frame_buttons, text="刷新文件列表", command=self.refresh_file_list, width=14).pack(fill="x", pady=3)
        ttk.Button(frame_buttons, text="清空列表", command=self.clear_list, width=14).pack(fill="x", pady=3)
        ttk.Button(frame_buttons, text="开始合并", command=self.start_merge_thread, width=14).pack(fill="x", pady=20)
        ttk.Button(frame_buttons, text="打印PDF", command=self.print_pdf, width=14).pack(fill="x", pady=3)
        ttk.Button(frame_buttons, text="退出", command=self.on_closing, width=14).pack(fill="x", pady=3)

        # 进度条与状态栏
        frame_bot = ttk.Frame(self, padding=pad)
        frame_bot.pack(fill="x")
        self.progress = ttk.Progressbar(frame_bot, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=6, pady=(0,6))
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(frame_bot, textvariable=self.status_var).pack(anchor="w", padx=6)

    def choose_input_dir(self):
        d = filedialog.askdirectory(title="选择包含 PDF 和图片的文件夹")
        if d:
            self.input_dir.set(d)
            self.refresh_file_list()

    def choose_output_file(self):
        f = filedialog.asksaveasfilename(title="保存合并的 PDF 为", defaultextension=".pdf", filetypes=[("PDF 文件","*.pdf")])
        if f:
            self.output_file.set(f)

    def is_image_file(self, path: Path) -> bool:
        """检查文件是否为支持的图片格式"""
        image_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff'}
        return path.suffix.lower() in image_extensions
    
    def collect_files(self, directory: Path, recursive: bool):
        """收集 PDF 和图片文件"""
        files = []
        if recursive:
            # 递归搜索
            pdfs = list(directory.glob("**/*.pdf"))
            images = []
            for ext in ['*.jpg', '*.jpeg', '*.png', '*.bmp', '*.tif', '*.tiff']:
                images.extend(directory.glob(f"**/{ext}"))
            files = pdfs + images
        else:
            # 只搜索当前目录
            pdfs = list(directory.glob("*.pdf"))
            images = []
            for ext in ['*.jpg', '*.jpeg', '*.png', '*.bmp', '*.tif', '*.tiff']:
                images.extend(directory.glob(ext))
            files = pdfs + images
        
        # 过滤出文件（排除目录）
        return [p for p in files if p.is_file()]

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
        files = self.collect_files(p, self.recursive.get())
        if not files:
            self.status_var.set("未找到 PDF 或图片文件。")
            return

        # 排序
        if self.sort_by.get() == "name":
            files.sort(key=lambda x: x.name.lower())
        else:
            files.sort(key=lambda x: x.stat().st_mtime)

        pdf_count = sum(1 for f in files if f.suffix.lower() == '.pdf')
        img_count = len(files) - pdf_count
        
        for f in files:
            # 在列表中显示文件类型标识
            if self.is_image_file(f):
                self.listbox.insert(tk.END, f"[图片] {f.name}")
            else:
                self.listbox.insert(tk.END, f"[PDF] {f.name}")
        self.file_list = files
        
        status_msg = f"找到 {len(self.file_list)} 个文件"
        if pdf_count > 0 and img_count > 0:
            status_msg += f"（{pdf_count} 个 PDF，{img_count} 个图片）"
        elif pdf_count > 0:
            status_msg += f"（{pdf_count} 个 PDF）"
        elif img_count > 0:
            status_msg += f"（{img_count} 个图片）"
        self.status_var.set(status_msg)

    def clear_list(self):
        self.listbox.delete(0, tk.END)
        self.file_list = []
        self.status_var.set("列表已清空。")

    def start_merge_thread(self):
        # 不要在主线程做合并，避免 GUI 假死
        if self.merge_thread and self.merge_thread.is_alive():
            messagebox.showwarning("正在合并", "合并操作正在进行中，请等待完成。")
            return
        self.merge_thread = threading.Thread(target=self.merge_pdfs, daemon=True)
        self.merge_thread.start()

    def resize_image_keep_ratio(self, img: Image.Image, max_width: int, max_height: int) -> Image.Image:
        """调整图片大小，保持宽高比，不超过最大尺寸"""
        img_width, img_height = img.size
        
        # 计算缩放比例，保持宽高比
        scale = min(max_width / img_width, max_height / img_height)
        new_width = int(img_width * scale)
        new_height = int(img_height * scale)
        
        # 调整图片大小
        img_resized = img.resize((new_width, new_height), RESAMPLE_MODE)
        return img_resized

    def create_page_with_items(self, items: list) -> Image.Image:
        """将多个图片排列到一张A4纸上（每页最多2个）
        
        Args:
            items: PIL Image 对象列表，最多2个
        
        Returns:
            包含所有图片的A4尺寸图片
        """
        target_width, target_height = STANDARD_PAGE_SIZE
        page = Image.new('RGB', STANDARD_PAGE_SIZE, (255, 255, 255))
        
        if len(items) == 0:
            return page
        
        # 统一布局：纵向 2 行（1 张也使用同样的布局）
        v_margin = int(target_height * 0.03)  # 垂直边距
        h_margin = int(target_width * 0.05)  # 水平边距

        cell_width = target_width - 2 * h_margin
        cell_height = int((target_height - 3 * v_margin) / 2)

        for i, img in enumerate(items[:2]):
            resized = self.resize_image_keep_ratio(img, cell_width, cell_height)
            x = h_margin + (cell_width - resized.width) // 2
            y = v_margin + i * (cell_height + v_margin) + (cell_height - resized.height) // 2
            page.paste(resized, (x, y))
        
        return page

    def image_to_pil(self, image_path: Path) -> Image.Image:
        """将图片文件转换为 PIL Image 对象（保持宽高比）"""
        if not PIL_AVAILABLE:
            raise ImportError("需要安装 Pillow 库来处理图片文件：pip install Pillow")
        
        try:
            img = Image.open(str(image_path))
            # 如果是 RGBA 模式，转换为 RGB（PDF 不支持透明度）
            if img.mode == 'RGBA':
                rgb_img = Image.new('RGB', img.size, (255, 255, 255))
                rgb_img.paste(img, mask=img.split()[3])  # 使用 alpha 通道作为 mask
                img = rgb_img
            elif img.mode not in ('RGB', 'L'):
                img = img.convert('RGB')
            return img
        except Exception as e:
            raise Exception(f"图片加载失败：{e}")

    def pdf_to_pil_images(self, pdf_path: Path) -> list:
        """将 PDF 的每一页转换为 PIL Image 对象列表（保持宽高比）"""
        if not PIL_AVAILABLE:
            raise ImportError("需要安装 Pillow 库来处理图片文件：pip install Pillow")
        
        if not FITZ_AVAILABLE:
            raise ImportError("需要安装 PyMuPDF 库来处理 PDF 文件：pip install PyMuPDF")
        
        try:
            # 使用 PyMuPDF 打开 PDF
            pdf_doc = fitz.open(str(pdf_path))
            images = []
            
            # 将每一页转换为图片
            for page_num in range(len(pdf_doc)):
                page = pdf_doc[page_num]
                # 渲染为图片，使用高分辨率（300 DPI）
                zoom = 300 / 72  # 300 DPI
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)
                
                # 转换为 PIL Image
                img_data = pix.tobytes("ppm")
                img = Image.open(BytesIO(img_data))
                images.append(img)
            
            pdf_doc.close()
            
            if not images:
                raise Exception("PDF 文件没有页面")
            
            return images
        except Exception as e:
            raise Exception(f"PDF 转换失败：{e}")

    def merge_pdfs(self):
        try:
            # 校验
            if not self.file_list:
                # 如果列表为空尝试自动刷新
                self.refresh_file_list()
                if not self.file_list:
                    messagebox.showwarning("没有文件", "未找到待合并的 PDF 或图片文件。")
                    return

            # 检查是否有图片文件需要 Pillow
            has_images = any(self.is_image_file(f) for f in self.file_list)
            has_pdfs = any(not self.is_image_file(f) for f in self.file_list)
            
            if has_images and not PIL_AVAILABLE:
                messagebox.showerror("缺少依赖", "检测到图片文件，但未安装 Pillow 库。\n\n请运行：pip install Pillow")
                return
            
            # 检查是否有 PDF 文件需要 PyMuPDF（用于调整大小）
            if has_pdfs and not FITZ_AVAILABLE:
                messagebox.showerror("缺少依赖", "检测到 PDF 文件，需要 PyMuPDF 库来调整 PDF 页面大小。\n\n请运行：pip install PyMuPDF")
                return

            out_path = self.output_file.get().strip()
            if not out_path:
                messagebox.showwarning("输出路径", "请先选择输出文件路径。")
                return
            out_path = Path(out_path)
            out_path.parent.mkdir(parents=True, exist_ok=True)

            # 第一步：将所有文件转换为 PIL Image 列表
            self.progress["maximum"] = len(self.file_list) + 100  # 预留空间用于组合页面
            self.progress["value"] = 0
            self.status_var.set("正在转换文件...")
            
            all_images = []  # 存储所有转换后的图片
            processed_count = 0
            skipped_count = 0
            
            for idx, file_path in enumerate(self.file_list, start=1):
                self.status_var.set(f"转换中：{file_path.name} （{idx}/{len(self.file_list)}）")
                self.update_idletasks()
                try:
                    if self.is_image_file(file_path):
                        # 处理图片文件：转换为 PIL Image
                        try:
                            img = self.image_to_pil(file_path)
                            all_images.append(img)
                            processed_count += 1
                            self.append_log(f"已转换：{file_path.name}")
                        except Exception as e:
                            self.append_log(f"图片转换失败（跳过）：{file_path.name}，原因：{e}")
                            skipped_count += 1
                    else:
                        # 处理 PDF 文件：转换为 PIL Image 列表
                        try:
                            # 先检查是否加密
                            reader = PdfReader(str(file_path))
                            if getattr(reader, "is_encrypted", False):
                                if self.skip_encrypted.get():
                                    self.listbox.itemconfig(idx-1, fg="gray")
                                    self.append_log(f"跳过加密：{file_path.name}")
                                    skipped_count += 1
                                    self.progress["value"] = idx
                                    continue
                                else:
                                    try:
                                        reader.decrypt("")  # 尝试空密码
                                    except Exception:
                                        self.append_log(f"无法解密，跳过：{file_path.name}")
                                        skipped_count += 1
                                        self.progress["value"] = idx
                                        continue
                            
                            # 将 PDF 的每一页转换为图片
                            pdf_images = self.pdf_to_pil_images(file_path)
                            all_images.extend(pdf_images)
                            processed_count += len(pdf_images)
                            self.append_log(f"已转换：{file_path.name}（{len(pdf_images)} 页）")
                        except Exception as e:
                            self.append_log(f"PDF 处理失败（跳过）：{file_path.name}，原因：{e}")
                            skipped_count += 1
                except Exception as e:
                    self.append_log(f"处理失败（跳过）：{file_path.name}，原因：{e}")
                    skipped_count += 1
                finally:
                    self.progress["value"] = idx
                    self.update_idletasks()
                    time.sleep(0.01)
            
            if not all_images:
                messagebox.showerror("失败", "没有成功转换任何文件，未生成输出文件。")
                self.status_var.set("合并失败：没有转换任何文件。")
                return
            
            # 第二步：将图片分组，每3个组合到一张A4纸上
            self.status_var.set("正在组合页面...")
            pages = []  # 存储组合后的A4页面
            
            items_per_page = 2
            total_pages = (len(all_images) + items_per_page - 1) // items_per_page  # 向上取整
            self.progress["maximum"] = len(self.file_list) + total_pages
            
            for i in range(0, len(all_images), items_per_page):
                page_images = all_images[i:i + items_per_page]
                page = self.create_page_with_items(page_images)
                pages.append(page)
                self.progress["value"] = len(self.file_list) + len(pages)
                self.status_var.set(f"已组合 {len(pages)}/{total_pages} 页...")
                self.update_idletasks()
            
            # 第三步：将所有页面保存为PDF
            self.status_var.set("正在生成PDF...")
            if not pages:
                messagebox.showerror("失败", "没有生成任何页面。")
                self.status_var.set("合并失败：没有生成页面。")
                return
            
            pdf_bytes = BytesIO()
            pages[0].save(
                pdf_bytes,
                format='PDF',
                resolution=300.0,
                save_all=True,
                append_images=pages[1:] if len(pages) > 1 else []
            )
            pdf_bytes.seek(0)
            
            # 写入最终文件
            with open(str(out_path), 'wb') as f:
                f.write(pdf_bytes.read())
            
            self.last_merged_pdf = out_path
            
            # 统计信息
            total_items = len(all_images)
            total_pages = len(pages)
            result_msg = f"合并完成！\n\n"
            result_msg += f"共处理 {processed_count} 个文件项"
            if skipped_count > 0:
                result_msg += f"，跳过 {skipped_count} 个文件"
            result_msg += f"\n共 {total_items} 个图片/PDF页面"
            result_msg += f"\n组合为 {total_pages} 张A4页面（每页最多2个）"
            result_msg += f"\n\n输出文件：{out_path}"
            
            messagebox.showinfo("完成", result_msg)
            self.status_var.set(f"合并完成：{total_pages} 页")
        except Exception as ex:
            traceback.print_exc()
            messagebox.showerror("错误", f"合并过程中发生错误：\n{ex}")
            self.status_var.set("合并出错。")

    def append_log(self, text):
        # 将 log 显示在 listbox 的下方（直接在 listbox 中标注）
        # 这里不新增 UI 元素，仅在控制台打印，需更多可把日志放到 Text 控件里
        print(text)

    def print_pdf(self):
        """打印合并后的PDF文件"""
        # 优先使用最后合并的PDF，如果没有则使用输出文件路径
        pdf_path = self.last_merged_pdf
        if not pdf_path:
            out_path = self.output_file.get().strip()
            if not out_path:
                messagebox.showwarning("没有PDF", "请先合并PDF文件或选择输出文件路径。")
                return
            pdf_path = Path(out_path)
        
        if isinstance(pdf_path, str):
            pdf_path = Path(pdf_path)
        
        if not pdf_path.exists():
            messagebox.showerror("文件不存在", f"PDF文件不存在：\n{pdf_path}\n\n请先完成合并操作。")
            self.status_var.set("打印失败：文件不存在。")
            return
        
        try:
            system = platform.system()
            pdf_str = str(pdf_path.absolute())
            
            if system == "Windows":
                # Windows: 尝试多种打印方法
                printed = False
                # 方法1: 尝试使用 os.startfile 的 print 操作
                try:
                    os.startfile(pdf_str, "print")
                    printed = True
                    self.status_var.set(f"正在打印：{pdf_path.name}")
                except (OSError, ValueError):
                    # 方法2: 使用 PowerShell 打印
                    try:
                        # 使用 PowerShell 的 Start-Process 命令打印
                        ps_command = f'Start-Process -FilePath "{pdf_str}" -Verb Print'
                        subprocess.run(
                            ["powershell", "-Command", ps_command],
                            check=True,
                            timeout=10,
                            creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
                        )
                        printed = True
                        self.status_var.set(f"已发送打印任务：{pdf_path.name}")
                    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, FileNotFoundError):
                        # 方法3: 打开文件让用户手动打印
                        try:
                            os.startfile(pdf_str)  # 不带 print 参数，直接打开
                            messagebox.showinfo("提示", f"已打开PDF文件，请使用 Ctrl+P 或文件菜单中的打印功能进行打印。\n\n文件：{pdf_path.name}")
                            self.status_var.set(f"已打开文件：{pdf_path.name}，请手动打印")
                            printed = True
                        except Exception:
                            pass
                
                if not printed:
                    raise Exception("无法执行打印操作，请检查系统是否安装了PDF查看器。")
            elif system == "Darwin":  # macOS
                # macOS: 使用 lp 命令
                subprocess.run(["lp", pdf_str], check=True)
                self.status_var.set(f"已发送到打印机：{pdf_path.name}")
            else:  # Linux
                # Linux: 尝试使用 lp 或 lpr 命令
                try:
                    subprocess.run(["lp", pdf_str], check=True)
                except (subprocess.CalledProcessError, FileNotFoundError):
                    try:
                        subprocess.run(["lpr", pdf_str], check=True)
                    except (subprocess.CalledProcessError, FileNotFoundError):
                        # 如果都没有，尝试用默认程序打开
                        subprocess.run(["xdg-open", pdf_str], check=True)
                        messagebox.showinfo("提示", "已用默认程序打开PDF，请手动打印。")
                self.status_var.set(f"已发送到打印机：{pdf_path.name}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("打印错误", f"打印过程中发生错误：\n{e}")
            self.status_var.set("打印出错。")

    def on_closing(self):
        """处理窗口关闭事件"""
        # 检查是否有正在运行的合并操作
        if self.merge_thread and self.merge_thread.is_alive():
            if messagebox.askokcancel("退出", "合并操作正在进行中，确定要退出吗？"):
                self.destroy()
                sys.exit(0)
        else:
            self.destroy()
            sys.exit(0)

def main():
    app = PDFMergerGUI()
    # 如果启动时需要设置默认目录/输出，可在这里赋值，例如：
    # app.input_dir.set(r"C:\pdfs")
    # app.output_file.set(r"C:\out\merged.pdf")
    app.mainloop()

if __name__ == "__main__":
    main()
