import re
from collections import defaultdict, Counter
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from typing import Optional, Tuple

# ----------------------------
# 参数（把下面变量改为你的实际路径 / sheet 名称等）
path_a = r"C:\Users\Alan\Desktop\润彩金顶街10.xlsx"        # A 表格路径
path_b = r"C:\Users\Alan\xwechat_files\wxid_vntsv125k1tx22_51c1\msg\file\2025-10\北京诚善堂金顶街药业有限公司.xlsx"        # B 表格路径
path_a_out = r"E:\项目\诚善堂\对账信息\润彩金顶街\润彩金顶街10.xlsx" # 标红后保存的 A 表新路径
path_b_out = r"E:\项目\诚善堂\对账信息\润彩金顶街\北京诚善堂金顶街药业有限公司.xlsx" # 标红后保存的 B 表新路径

sheet_a_name: Optional[str] = None  # None -> 使用第一个 sheet
sheet_b_name: Optional[str] = None

start_row = 2       # 从第几行开始读取（你的描述是从第2行）
col_name_a = "G"    # A 表品名列字母
col_qty_a  = "N"    # A 表进货数量列字母

col_name_b = "E"    # B 表品名列字母
col_qty_b  = "F"    # B 表进货数量列字母

# ------------- 品名等价规则（可扩展/修改） -------------
# 优先级：explicit_map -> regex_replacements -> normalize_name 的默认处理
# explicit_map：完全匹配（规范化后）的显式映射（适合少量特殊短语）


# regex_replacements：按正则替换规范化名称（可以把括号、大小写差异处理外的其他变体统一）
# 每项是 (pattern, replacement) ，会在 normalize 后依次应用
regex_replacements = [
    # 把 "血糖试条(葡萄糖脱氢酶法)" 或包含该文案的变体都替换为 "血糖试条"
    (r"血糖试条.*葡萄糖脱氢酶法.*", "血糖试条"),
    (r"血糖试条\(.*葡萄糖脱氢酶法.*\)", "血糖试条"),
    # 把大小写 e/E 可能带来的差异通过 lower 处理即可，但这里可以处理 "vitamin e" 类
    # 合并品牌与品名顺序（把“冈本...”的多种顺序换成统一）
    (r"(^|.*)冈本(.*天然胶乳橡胶避孕套)(.*)", "冈本天然胶乳橡胶避孕套"),
    (r"(天然胶乳橡胶避孕套)(.*冈本)(.*)", "冈本天然胶乳橡胶避孕套"),
    # 你可以在这里加入更多自定义替换规则
]
# -------------------------------------------------------

# 正则：移除括号及括号内内容（包含中英文全角括号）
_BRACKET_PATTERN = re.compile(r"[\(\[{\（\【][^\)\]\}】\）}]*[\)\]\}】\）}]")


def normalize_name(s: Optional[str]) -> str:
    """去括号、合并空格、lower、应用正则替换与显式映射，返回规范名（小写）。"""
    if s is None:
        return ""
    s = str(s)
    s = _BRACKET_PATTERN.sub("", s)
    s = s.replace("\u3000", " ").strip()
    s = re.sub(r"\s+", " ", s)
    s = s.lower()
    s = s.strip(" -–—_：:，,。.；;\"'“”")
    for pat, repl in regex_replacements:
        s = re.sub(pat, repl, s)
    if s in explicit_map:
        return explicit_map[s].lower()
    return s

# explicit_map 初始字典（可为空）
explicit_map = {}

# add_equivalence 必须放在 normalize_name 定义之后
def add_equivalence(raw_variant: str, canonical: str, explicit_map: dict):
    """
    将原始变体（raw_variant）规范化后加入 explicit_map，使其映射到 canonical（统一名称）。
    返回规范化后的键，便于调试。
    """
    # 使用已经定义好的 normalize_name
    norm = normalize_name(raw_variant)
    explicit_map[norm] = canonical.lower()
    return norm

# 示例：把新规则加入 explicit_map
add_equivalence("对乙酰氨基酚混悬液", "对乙酰氨基酚口服混悬液", explicit_map)
add_equivalence("雅嘉莱茉莉维生素E乳", "雅嘉莱茉莉维生素e乳", explicit_map)
add_equivalence("血糖试条(葡萄糖脱氢酶法)", "血糖试条", explicit_map)
add_equivalence("血糖试条", "血糖试条", explicit_map)
add_equivalence("天然胶乳橡胶避孕套冈本避孕套", "冈本天然胶乳橡胶避孕套", explicit_map)
add_equivalence("冈本天然胶乳橡胶避孕套", "冈本天然胶乳橡胶避孕套", explicit_map)
add_equivalence("脱脂棉球", "欧洁脱脂棉球", explicit_map)
add_equivalence("什果冰润唇膏(香橙+蜜柑/草莓)", "什果冰润唇膏-香橙+蜜柑", explicit_map)
add_equivalence("蜜炼川贝枇杷膏", "京都念慈菴蜜炼川贝枇杷膏", explicit_map)
add_equivalence("碘伏消毒液", "利尔康牌碘伏消毒液", explicit_map)
add_equivalence("美敏伪麻溶液", "美敏伪麻口服溶液", explicit_map)
add_equivalence("云南白药粉", "云南白药", explicit_map)
add_equivalence("怡成血糖仪", "血糖仪+血糖试条(葡萄糖脱氢酶法)+一次性使用末梢采血针+采血笔", explicit_map)
add_equivalence("云南白药牙膏(薄荷清爽)", "云南白药牙膏薄荷清爽香型", explicit_map)
add_equivalence("云南白药牙膏", "云南白药牙膏薄荷清爽香型", explicit_map)
def parse_qty_key(val) -> Tuple[str, object]:
    """数量优先解析为数字（int/float），否则用规范化字符串表示。"""
    if val is None:
        return ('str', "")
    s = str(val).strip()
    s_clean = s.replace(",", "").replace("，", "").replace(" ", "").replace("\u3000", "")
    if s_clean == "":
        return ('str', "")
    try:
        f = float(s_clean)
        if abs(f - int(f)) < 1e-9:
            return ('num', int(round(f)))
        else:
            return ('num', round(f, 6))
    except Exception:
        return ('str', s.lower().strip())

def col_to_index(col: str) -> int:
    col = col.upper()
    total = 0
    exp = 0
    for ch in reversed(col):
        total += (ord(ch) - ord('A') + 1) * (26 ** exp)
        exp += 1
    return total

def load_ws(path: str, sheet_name: Optional[str]):
    wb = load_workbook(path, read_only=False, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = wb[wb.sheetnames[0]]
    return wb, ws

def build_entries(ws, name_col: str, qty_col: str, start_row: int):
    """
    返回 dict: normalized_name -> list of entry dicts
    entry: {'row': int, 'raw_name': str, 'raw_qty': original_value, 'qty_key': ('num'/'str', value)}
    """
    m = defaultdict(list)
    max_r = ws.max_row if ws.max_row is not None else 0
    for r in range(start_row, max_r + 1):
        raw_name = ws.cell(row=r, column=col_to_index(name_col)).value
        raw_qty  = ws.cell(row=r, column=col_to_index(qty_col)).value
        normalized = normalize_name(raw_name)
        if normalized == "":
            continue
        qty_key = parse_qty_key(raw_qty)
        m[normalized].append({
            'row': r,
            'raw_name': "" if raw_name is None else str(raw_name),
            'raw_qty': raw_qty,
            'qty_key': qty_key
        })
    return m

def compare_and_mark(path_a, path_b,
                     sheet_a_name, sheet_b_name,
                     start_row,
                     col_name_a, col_qty_a,
                     col_name_b, col_qty_b,
                     path_a_out, path_b_out):
    # 加载工作表（非只读）以便后续保存标红
    wb_a, ws_a = load_ws(path_a, sheet_a_name)
    wb_b, ws_b = load_ws(path_b, sheet_b_name)

    a_map = build_entries(ws_a, col_name_a, col_qty_a, start_row)
    b_map = build_entries(ws_b, col_name_b, col_qty_b, start_row)

    all_names = set(a_map.keys()) | set(b_map.keys())
    total_unmatched = 0

    unmatched_rows_a = set()
    unmatched_rows_b = set()

    for name in sorted(all_names):
        list_a = a_map.get(name, [])
        list_b = b_map.get(name, [])

        # multiset 比较（Counter）
        counter_a = Counter([tuple(e['qty_key']) for e in list_a])
        counter_b = Counter([tuple(e['qty_key']) for e in list_b])

        if counter_a == counter_b:
            continue  # 完全匹配，跳过

        # 具体未匹配项：从 B 建桶并逐个消耗匹配项
        b_bucket = defaultdict(list)
        for e in list_b:
            b_bucket[tuple(e['qty_key'])].append(e)

        unmatched_a = []
        for ea in list_a:
            k = tuple(ea['qty_key'])
            if b_bucket.get(k):
                b_bucket[k].pop()
                if not b_bucket[k]:
                    del b_bucket[k]
            else:
                unmatched_a.append(ea)

        unmatched_b = []
        for lst in b_bucket.values():
            unmatched_b.extend(lst)

        # 汇总未匹配行号（用于标红）
        for e in unmatched_a:
            unmatched_rows_a.add(e['row'])
        for e in unmatched_b:
            unmatched_rows_b.add(e['row'])

        # 输出控制台明细
        total_unmatched += len(unmatched_a) + len(unmatched_b)
        print("="*80)
        print(f"规范名: '{name}'  （A共有 {len(list_a)} 行，B共有 {len(list_b)} 行）")
        if unmatched_a:
            print("  A 表未配对：")
            for e in unmatched_a:
                print(f"    A 行 {e['row']}: 原名='{e['raw_name']}' ，进货数量='{e['raw_qty']}'")
        else:
            print("  A 表未配对：无")
        if unmatched_b:
            print("  B 表未配对：")
            for e in unmatched_b:
                print(f"    B 行 {e['row']}: 原名='{e['raw_name']}' ，进货数量='{e['raw_qty']}'")
        else:
            print("  B 表未配对：无")

    # 控制台总结
    if total_unmatched == 0:
        print("比较完成：未发现不匹配（按规范名和数量 multiset 比较）。")
    else:
        print("="*80)
        print(f"比较完成：共发现 {total_unmatched} 条未配对记录（在控制台已列出详情）。")

    # 标红并保存 A、B 到指定新路径
    red_fill = PatternFill(start_color="F08080", end_color="F08080", fill_type="solid")

    # 标红 A
    if unmatched_rows_a:
        max_col_a = ws_a.max_column if ws_a.max_column is not None else col_to_index(col_qty_a)
        for r in unmatched_rows_a:
            # 标注整行（从第1列到 max_col_a）
            for c in range(1, max_col_a + 1):
                cell = ws_a.cell(row=r, column=c)
                cell.fill = red_fill
    # 标红 B
    if unmatched_rows_b:
        max_col_b = ws_b.max_column if ws_b.max_column is not None else col_to_index(col_qty_b)
        for r in unmatched_rows_b:
            for c in range(1, max_col_b + 1):
                cell = ws_b.cell(row=r, column=c)
                cell.fill = red_fill

    # 保存到新文件（如果路径与原路径相同会覆盖原文件）
    wb_a.save(path_a_out)
    wb_b.save(path_b_out)
    print(f"已将 A 表保存为: {path_a_out} （未匹配行已标红）")
    print(f"已将 B 表保存为: {path_b_out} （未匹配行已标红）")

if __name__ == "__main__":
    compare_and_mark(
        path_a=path_a,
        path_b=path_b,
        sheet_a_name=sheet_a_name,
        sheet_b_name=sheet_b_name,
        start_row=start_row,
        col_name_a=col_name_a,
        col_qty_a=col_qty_a,
        col_name_b=col_name_b,
        col_qty_b=col_qty_b,
        path_a_out=path_a_out,
        path_b_out=path_b_out
    )