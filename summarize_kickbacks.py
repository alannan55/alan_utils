import os.path
import re
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import numbers


def none2zero(amount):
    if not amount:
        amount = 0
    return amount

def append_totals_zip(new_ws, exclude_cols=None, header_row=1, min_col=1, start_row=2):
    """
    在 new_ws 的最后追加一行 ['总计', sum_col2, sum_col3, ...]
    exclude_cols: None 或 可迭代（int 列号(1-based) 或 列字母 'B' 或 表头文本 '姓名'）
    header_row: 用于按表头名排除列的行号（默认第1行）
    min_col: 读取时的起始列（默认 1）
    start_row: 数据起始行（默认第2行）
    """
    max_col = new_ws.max_column
    # 没有数据列
    if max_col < min_col:
        new_ws.append(["总计"])
        return

    # 读取行数据（保证每行长度等于 max_col-min_col+1）
    rows = list(new_ws.iter_rows(min_row=start_row, max_row=new_ws.max_row,
                                 min_col=min_col, max_col=max_col,
                                 values_only=True))
    if not rows:
        new_ws.append(["总计"])
        return

    # 解析 exclude_cols 到列号集合（1-based 全表列号）
    exclude_set = set()
    if exclude_cols:
        for item in exclude_cols:
            if item is None:
                continue
            # 整数：直接当作列号
            if isinstance(item, int):
                exclude_set.add(item)
                continue
            # 字符串：尝试当作列字母
            if isinstance(item, str):
                s = item.strip()
                if not s:
                    continue
                # 尝试列字母 -> 列号
                try:
                    col_idx = column_index_from_string(s.upper())
                    exclude_set.add(col_idx)
                    continue
                except Exception:
                    pass
                # 否则当作表头文字，去 header_row 查找匹配列
                try:
                    header_cells = list(new_ws.iter_rows(min_row=header_row, max_row=header_row,
                                                         min_col=min_col, max_col=max_col,
                                                         values_only=False))[0]
                except Exception:
                    header_cells = []
                for i, cell in enumerate(header_cells):
                    # cell 的全表列号：
                    col_no = min_col + i
                    if cell.value == s:
                        exclude_set.add(col_no)

    # cols 是从 min_col 到 max_col 的列集合（每个 col 是从 start_row 开始的列值元组）
    cols = list(zip(*rows))
    sums = []
    # 对于 cols 的索引 i 对应全表列号 = min_col + i
    for i, col in enumerate(cols):
        col_no = min_col + i
        # 按你原来逻辑，第一列通常不求和；如果 user 也把第一列列入 exclude_cols，仍然会被排除
        if i == 0 or (col_no in exclude_set):
            # 在结果里保留占位，确保列对齐（与原行为一致：忽略第一列）
            # 对第一列仍保留占位（第一列位置用 '总计' 代替，因此这里填 ''）
            sums.append('')
            continue

        total = 0.0
        for v in col:
            if v is None or v == "":
                continue
            if isinstance(v, numbers.Number):
                total += v
            else:
                try:
                    total += float(v)
                except Exception:
                    # 非数值则忽略
                    continue
        # 如果 total 是整数，转成 int 显示更漂亮
        if abs(total - round(total)) < 1e-9:
            total = int(round(total))
        sums.append(total)

    # sums 对应从 min_col 到 max_col 的每一列，这里构造要 append 的行：
    # 但你原版本是 ["总计"] + sums_for_cols_from_2_onwards
    # 由于 sums 包含了第一列的占位，直接这样拼接即可：
    new_row = ["总计"] + sums[1:]  # 跳过 sums[0] 因为第一个位置为第一列，占位已由 "总计" 替代
    new_ws.append(new_row)

def sanitize_sheet_title(raw_title: str, max_len: int = 31) -> str:
    """
    将 raw_title 转为合法的 Excel sheet 名称：
    - 替换非法字符 [: \\ / ? * [ ] ] 为下划线
    - 去首尾空白
    - 截断到 max_len 字符
    返回不会为空的字符串（空则返回 'Sheet'）
    """
    if raw_title is None:
        raw_title = ""
    # Excel 不允许的字符集合
    invalid_chars = r'[:\\\/\?\*\[\]]'
    title = re.sub(invalid_chars, '_', str(raw_title))
    title = title.strip()
    if not title:
        title = "Sheet"
    if len(title) > max_len:
        title = title[:max_len]
    return title

def copy_rows(sales_index, sheet, new_wb):
    """
    将 sales_index 指定的行复制到 new_wb 中对应的 sheet（sheet 名用 sanitize_sheet_title 处理）。
    - 若目标 sheet 已存在，则直接追加（不删除已有总计行）。
    - 若目标 sheet 不存在或为空，则写入表头到第1、第2行（直接用 cell 赋值，避免顶端空行）。
    - 每次追加后调用 append_totals_zip 追加新的总计行（与你的要求一致）。
    """
    for salesman, index_list in sales_index.items():
        if not salesman or not index_list:
            continue

        safe_title = sanitize_sheet_title(salesman)

        # 复用已存在 sheet（若存在），否则新建
        if safe_title in new_wb.sheetnames:
            cur_ws = new_wb[safe_title]
        else:
            cur_ws = new_wb.create_sheet(title=safe_title)

        # 写第1行（如果源表有第1行）
        try:
            header1 = [cell.value for cell in sheet[1]]
            cur_ws.append(header1)
            # for col_idx, val in enumerate(header1, start=1):
            #     cur_ws.cell(row=1, column=col_idx, value=val)
        except Exception:
            pass

        # 写第2行（如果源表有第2行）
        try:
            header2 = [cell.value for cell in sheet[2]]
            cur_ws.append(header2)
            # for col_idx, val in enumerate(header2, start=1):
            #     cur_ws.cell(row=2, column=col_idx, value=val)
        except Exception:
            pass

        # 追加索引指定的行（保持你原来的 index+1 偏移）
        for index in index_list:
            row_values = [cell.value for cell in sheet[index+1]]
            cur_ws.append(row_values)

        # 追加新的总计行（不删原有总计）
        append_totals_zip(cur_ws, exclude_cols=['A', 'B', 'C', 'D', 'E', 'F', 'G', 'R', 'S'])
        cur_ws.append([""])


def summarize(input_path, save_wb):
    wb = load_workbook(input_path, data_only=True)
    basename, suffix = os.path.splitext(input_path)

    new_ws = save_wb.active
    new_ws.title = "汇总"

    # read
    sheets = wb.worksheets
    sheet = sheets[0]
    sales_index = {}

    for i in range(2, len(sheet['F'])):
        cur_salesman = sheet['F'][i].value
        if cur_salesman:
            if "王晨宇" in cur_salesman:
                cur_salesman = "王晨宇"
            if cur_salesman != "总计":
                if cur_salesman not in sales_index.keys():
                    sales_index[cur_salesman] = [i]
                else:
                    sales_index[cur_salesman].append(i)

    copy_rows(sales_index, sheet, save_wb)

    header1 = [cell.value for cell in sheet[1]]
    new_ws.append(header1)
    new_ws.append([""])
    name_a, name_b, name_c, name_d = sheet['J'][1].value, sheet['L'][1].value, sheet['N'][1].value, sheet['P'][1].value
    new_ws['A2'] = '业务员'
    new_ws['B2'] = '总数量'
    new_ws['C2'] = '总金额'
    new_ws['D2'] = name_a
    new_ws['E2'] = name_a + '金额'
    new_ws['F2'] = name_b
    new_ws['G2'] = name_b + '金额'
    new_ws['H2'] = name_c
    new_ws['I2'] = name_c + '金额'
    new_ws['J2'] = name_d
    new_ws['K2'] = name_d + '金额'
    for salesman, index_list in sales_index.items():
        cur_sales_counts = {name_a: 0, name_b: 0, name_c: 0, name_d: 0}
        cur_sales_amounts = {name_a: 0, name_b: 0, name_c: 0, name_d: 0}
        cur_sales_total_counts = 0
        cur_sales_total_amounts = 0
        for i in index_list:
            count_a, count_b, count_c, count_d = sheet['J'][i].value, sheet['L'][i].value, sheet['N'][i].value, \
            sheet['P'][i].value
            count_a, count_b, count_c, count_d = none2zero(count_a), none2zero(count_b), none2zero(count_c), none2zero(
                count_d)
            amount_a, amount_b, amount_c, amount_d = sheet['K'][i].value, sheet['M'][i].value, sheet['O'][i].value, \
            sheet['Q'][i].value
            amount_a, amount_b, amount_c, amount_d = none2zero(amount_a), none2zero(amount_b), none2zero(
                amount_c), none2zero(amount_d)

            cur_sales_counts[name_a] += count_a
            cur_sales_counts[name_b] += count_b
            cur_sales_counts[name_c] += count_c
            cur_sales_counts[name_d] += count_d

            cur_sales_amounts[name_a] += amount_a
            cur_sales_amounts[name_b] += amount_b
            cur_sales_amounts[name_c] += amount_c
            cur_sales_amounts[name_d] += amount_d

            cur_sales_total_counts += count_a + count_b + count_c + count_d
            cur_sales_total_amounts += amount_a + amount_b + amount_c + amount_d

        cur_row = [salesman, cur_sales_total_counts, cur_sales_total_amounts,
                   cur_sales_counts[name_a], cur_sales_amounts[name_a], cur_sales_counts[name_b],
                   cur_sales_amounts[name_b],
                   cur_sales_counts[name_c], cur_sales_amounts[name_c], cur_sales_counts[name_d],
                   cur_sales_amounts[name_d]]
        new_ws.append(cur_row)

    append_totals_zip(new_ws, exclude_cols=["L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"])


def main(jindingjie_path, laoshan_path):
    save_path = r"D:\项目\诚善堂\付款信息\诚善堂10月工资\new.xlsx"
    save_wb = Workbook()
    summarize(jindingjie_path, save_wb)
    summarize(laoshan_path, save_wb)
    save_wb.save(save_path)


if __name__ == '__main__':
    jindingjie_path = r"D:\项目\诚善堂\付款信息\诚善堂10月工资\诚善堂金顶街25年9月份代金品种1.xlsx"
    laoshan_path = r"D:\项目\诚善堂\付款信息\诚善堂10月工资\诚善堂老山25年9月份代金品种1.xlsx"
    main(jindingjie_path, laoshan_path)