# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import ctypes
import re
import sys
import time
import traceback
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pythoncom
    import win32clipboard
    import win32com.client
except ImportError as exc:
    pythoncom = None
    win32clipboard = None
    win32com = None
    WIN32_IMPORT_ERROR = exc
else:
    WIN32_IMPORT_ERROR = None

try:
    from PIL import ImageGrab
except ImportError as exc:
    ImageGrab = None
    PIL_IMPORT_ERROR = exc
else:
    PIL_IMPORT_ERROR = None


EXCEL_PATH_FILTER = "Excel 文件 (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|所有文件 (*.*)|*.*"
TEMP_SHEET_NAME = "__voucher_shot_tmp__"
TOTAL_TEXT = "总计"
HEADER_PERSON_TEXT = "业务员"
SPECIAL_NAMES = {
    "徐春杰",
    "苏晓悦",
    "郭（月底群里发）",
    "健兴张辉",
}

XL_SCREEN = 1
XL_BITMAP = 2
XL_SOLID = 1


@dataclass
class PendingTask:
    sheet_index: int
    sheet_name: str
    store_name: str
    person_name: str
    start_row: int
    end_row: int
    amount_text: str


def excel_color(hex_color: str) -> int:
    value = hex_color.strip().lstrip("#")
    red = int(value[0:2], 16)
    green = int(value[2:4], 16)
    blue = int(value[4:6], 16)
    return red + green * 256 + blue * 65536


GREEN_COLOR = excel_color("#92D050")
RED_COLOR = excel_color("#FF0000")


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = text.replace("(", "（").replace(")", "）")
    text = re.sub(r"\s+", "", text)
    return text


def to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("￥", "").replace("¥", "")
    try:
        return float(text)
    except ValueError:
        return None


def is_zero(value) -> bool:
    number = to_number(value)
    return number is not None and abs(number) < 0.0000001


def last_used_row(ws) -> int:
    used_range = ws.UsedRange
    return int(used_range.Row + used_range.Rows.Count - 1)


def set_fill(cell, color: int) -> None:
    cell.Interior.Pattern = XL_SOLID
    cell.Interior.Color = color


def find_next_total_row(total_rows: list[int], after_row: int) -> int | None:
    for row in total_rows:
        if row > after_row:
            return row
    return None


def store_short_name(ws) -> str:
    sheet_name = str(ws.Name).strip()
    title = str(ws.Cells(1, 1).Text).strip()
    source = f"{sheet_name} {title}"

    if "老山" in source:
        return "老山"
    if "金顶街" in source:
        return "金顶街"

    name = sheet_name
    name = re.sub(r"^诚善堂", "", name)
    name = re.sub(r"药业$", "", name)
    name = re.sub(r"\d+月.*$", "", name)
    return name.strip() or sheet_name


def segment_person_name(ws, start_row: int, end_row: int) -> str:
    names: list[str] = []
    for row in range(start_row, end_row + 1):
        raw_text = str(ws.Cells(row, 7).Text).strip()
        normalized = normalize_text(raw_text)
        if not normalized:
            continue
        if normalized in {TOTAL_TEXT, HEADER_PERSON_TEXT}:
            continue
        names.append(raw_text.strip())

    if not names:
        return f"rows_{start_row}-{end_row}"

    counts = Counter(names)
    first_index = {name: index for index, name in enumerate(names)}
    return max(counts, key=lambda name: (counts[name], -first_index[name]))


def safe_file_part(text: str) -> str:
    text = str(text).strip() or "blank"
    for char in r'<>:"/\|?*':
        text = text.replace(char, "_")
    text = re.sub(r"\s+", "", text)
    return text[:80] or "blank"


def unique_output_path(output_dir: Path, base_name: str, used_names: set[str]) -> Path:
    safe_base = safe_file_part(base_name)
    candidate = safe_base
    index = 2
    while candidate.lower() in used_names:
        candidate = f"{safe_base}_{index}"
        index += 1
    used_names.add(candidate.lower())
    return output_dir / f"{candidate}.png"


def clear_clipboard() -> None:
    if win32clipboard is None:
        return
    opened = False
    try:
        win32clipboard.OpenClipboard()
        opened = True
        win32clipboard.EmptyClipboard()
    except Exception:
        pass
    finally:
        if opened:
            try:
                win32clipboard.CloseClipboard()
            except Exception:
                pass


def clipboard_sequence() -> int:
    try:
        return int(ctypes.windll.user32.GetClipboardSequenceNumber())
    except Exception:
        return 0


def image_has_visible_content(image) -> bool:
    try:
        rgb_image = image.convert("RGB")
        colors = rgb_image.getcolors(maxcolors=1000000)
    except Exception:
        return True
    if not colors:
        return True
    dark_pixels = sum(count for count, color in colors if max(color) < 90)
    return dark_pixels > 20


def save_range_as_png(rng, output_path: Path) -> None:
    if ImageGrab is None:
        raise RuntimeError(f"Pillow 未安装，无法保存截图：{PIL_IMPORT_ERROR}")

    last_error: Exception | None = None

    for attempt in range(1, 8):
        clear_clipboard()
        time.sleep(0.15)
        before_sequence = clipboard_sequence()

        try:
            rng.Worksheet.Activate()
            rng.Select()
            rng.CopyPicture(Appearance=XL_SCREEN, Format=XL_BITMAP)
        except Exception as exc:
            last_error = exc
            time.sleep(0.2 * attempt)
            continue

        deadline = time.time() + 1.2 + attempt * 0.35
        while time.time() < deadline:
            if pythoncom is not None:
                pythoncom.PumpWaitingMessages()
            image = ImageGrab.grabclipboard()
            if image is not None and hasattr(image, "save") and clipboard_sequence() != before_sequence:
                if image_has_visible_content(image):
                    if output_path.exists():
                        output_path.unlink()
                    image.save(output_path, "PNG")
                    clear_clipboard()
                    time.sleep(0.15)
                    return
            time.sleep(0.12)

    if last_error is not None:
        raise RuntimeError(f"Excel 复制截图失败：{last_error}")
    raise RuntimeError("Excel 没有把截图放入剪贴板，请关闭占用剪贴板的截图或远程控制工具后重试。")


def remove_temp_sheet(workbook) -> None:
    app = workbook.Application
    old_alerts = app.DisplayAlerts
    app.DisplayAlerts = False
    try:
        for index in range(workbook.Worksheets.Count, 0, -1):
            sheet = workbook.Worksheets(index)
            if sheet.Name == TEMP_SHEET_NAME:
                sheet.Delete()
    finally:
        app.DisplayAlerts = old_alerts


def capture_source_segment_as_png(ws, task: PendingTask, output_path: Path) -> None:
    hidden_states: list[tuple[int, bool]] = []

    try:
        if task.start_row > 3:
            for row in range(3, task.start_row):
                hidden_states.append((row, bool(ws.Rows(row).Hidden)))
            ws.Rows(f"3:{task.start_row - 1}").Hidden = True

        rng = ws.Range(ws.Cells(1, 1), ws.Cells(task.end_row, 10))
        save_range_as_png(rng, output_path)
    finally:
        for row, was_hidden in hidden_states:
            ws.Rows(row).Hidden = was_hidden


def process_worksheet(ws, sheet_index: int, log: Callable[[str], None]) -> list[PendingTask]:
    max_row = last_used_row(ws)
    total_rows: list[int] = []
    special_rows: list[int] = []

    for row in range(1, max_row + 1):
        g_text = normalize_text(ws.Cells(row, 7).Value)
        if g_text == TOTAL_TEXT:
            total_rows.append(row)
        elif g_text in SPECIAL_NAMES:
            special_rows.append(row)

    special_target_rows: set[int] = set()
    for special_row in special_rows:
        target_row = find_next_total_row(total_rows, special_row)
        if target_row is None:
            log(f"{ws.Name}: 第 {special_row} 行没有找到后续总计行")
        else:
            special_target_rows.add(target_row)

    tasks: list[PendingTask] = []
    green_count = 0
    red_count = 0
    store_name = store_short_name(ws)

    for index, total_row in enumerate(total_rows):
        amount_cell = ws.Cells(total_row, 10)
        zero_total = is_zero(amount_cell.Value)
        special_target = total_row in special_target_rows

        if zero_total or special_target:
            set_fill(amount_cell, GREEN_COLOR)
            green_count += 1
            continue

        set_fill(amount_cell, RED_COLOR)
        red_count += 1
        start_row = 1 if index == 0 else total_rows[index - 1] + 1
        person_name = segment_person_name(ws, start_row, total_row)
        tasks.append(
            PendingTask(
                sheet_index=sheet_index,
                sheet_name=str(ws.Name),
                store_name=store_name,
                person_name=person_name,
                start_row=start_row,
                end_row=total_row,
                amount_text=str(amount_cell.Text),
            )
        )

    log(f"{ws.Name}: 总计行 {len(total_rows)} 个，绿色 {green_count} 个，红色/待核销 {red_count} 个")
    return tasks


def export_screenshots(workbook, tasks: list[PendingTask], output_dir: Path, log: Callable[[str], None]) -> int:
    if not tasks:
        log("没有待核销行，不需要生成截图")
        return 0

    app = workbook.Application
    app.ScreenUpdating = True
    remove_temp_sheet(workbook)

    used_names: set[str] = set()
    for index, task in enumerate(tasks, start=1):
        source_ws = workbook.Worksheets(task.sheet_index)
        base_name = f"{task.store_name}_{task.person_name}"
        output_path = unique_output_path(output_dir, base_name, used_names)
        capture_source_segment_as_png(source_ws, task, output_path)
        log(f"截图 {index}/{len(tasks)}：{output_path.name}")
    return len(tasks)


def ensure_runtime_dependencies() -> None:
    if WIN32_IMPORT_ERROR is not None:
        raise RuntimeError(f"pywin32 未安装，无法控制 Excel：{WIN32_IMPORT_ERROR}")
    if PIL_IMPORT_ERROR is not None:
        raise RuntimeError(f"Pillow 未安装，无法保存截图：{PIL_IMPORT_ERROR}")


def reconcile_workbook(workbook_path: str | Path, output_dir: str | Path, log: Callable[[str], None]) -> None:
    ensure_runtime_dependencies()

    workbook_path = Path(workbook_path).expanduser().resolve()
    output_dir = Path(output_dir).expanduser().resolve()

    if not workbook_path.is_file():
        raise FileNotFoundError(f"找不到表格：{workbook_path}")
    output_dir.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        log("启动 Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        log("打开表格...")
        workbook = excel.Workbooks.Open(str(workbook_path), 0, False)
        if workbook.ReadOnly:
            raise RuntimeError("表格以只读方式打开。请先关闭 Excel 中打开的这个文件，再重新运行。")
        if workbook.Worksheets.Count < 2:
            raise RuntimeError("表格少于两个 sheet，无法按前两个 sheet 处理。")

        tasks: list[PendingTask] = []
        for sheet_index in (1, 2):
            ws = workbook.Worksheets(sheet_index)
            log(f"处理 sheet {sheet_index}：{ws.Name}")
            tasks.extend(process_worksheet(ws, sheet_index, log))

        log(f"待核销截图任务：{len(tasks)} 个")
        count = export_screenshots(workbook, tasks, output_dir, log)

        remove_temp_sheet(workbook)
        log("原地保存表格...")
        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
        log(f"完成，已导出截图 {count} 张")
    except Exception:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        raise
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


class VoucherApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("代金核销截图工具")
        self.geometry("780x455")
        self.resizable(False, False)

        self.workbook_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="表格路径").grid(row=0, column=0, sticky=tk.W, pady=5)
        workbook_entry = ttk.Entry(frame, textvariable=self.workbook_var, width=76)
        workbook_entry.grid(row=0, column=1, sticky=tk.W, padx=8)
        ttk.Button(frame, text="浏览", command=self.choose_workbook).grid(row=0, column=2, sticky=tk.W)

        ttk.Label(frame, text="截图目录").grid(row=1, column=0, sticky=tk.W, pady=5)
        output_entry = ttk.Entry(frame, textvariable=self.output_var, width=76)
        output_entry.grid(row=1, column=1, sticky=tk.W, padx=8)
        ttk.Button(frame, text="浏览", command=self.choose_output_dir).grid(row=1, column=2, sticky=tk.W)

        self.run_button = ttk.Button(frame, text="开始处理", command=self.run)
        self.run_button.grid(row=2, column=2, sticky=tk.E, pady=12)

        self.log_text = tk.Text(frame, height=16, width=104, state=tk.DISABLED)
        self.log_text.grid(row=3, column=0, columnspan=3, sticky=tk.NSEW, pady=(5, 0))

    def choose_workbook(self) -> None:
        path = filedialog.askopenfilename(title="选择表格", filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")])
        if path:
            self.workbook_var.set(path)

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择截图保存目录")
        if path:
            self.output_var.set(path)

    def log(self, message: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.update_idletasks()

    def run(self) -> None:
        workbook_path = self.workbook_var.get().strip()
        output_dir = self.output_var.get().strip()
        if not workbook_path or not output_dir:
            messagebox.showwarning("缺少路径", "请先选择表格路径和截图保存目录。")
            return

        ok = messagebox.askokcancel("确认原地保存", "程序会修改所选表格并原地保存。\n\n测试时请先使用副本。是否继续？")
        if not ok:
            return

        self.run_button.configure(state=tk.DISABLED)
        self.configure(cursor="watch")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)

        try:
            reconcile_workbook(workbook_path, output_dir, self.log)
            messagebox.showinfo("完成", "处理完成。")
        except Exception as exc:
            self.log("错误：" + str(exc))
            self.log(traceback.format_exc())
            messagebox.showerror("错误", str(exc))
        finally:
            self.configure(cursor="")
            self.run_button.configure(state=tk.NORMAL)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="代金核销标色并导出 Excel 样式截图")
    parser.add_argument("--run", action="store_true", help="不打开 UI，直接处理")
    parser.add_argument("--workbook", help="Excel 表格路径")
    parser.add_argument("--output", help="截图保存目录")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    if args.run:
        if not args.workbook or not args.output:
            print("--run 需要同时提供 --workbook 和 --output", file=sys.stderr)
            return 2
        reconcile_workbook(args.workbook, args.output, lambda message: print(f"[{time.strftime('%H:%M:%S')}] {message}"))
        return 0

    app = VoucherApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
