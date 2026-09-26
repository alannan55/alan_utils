# -*- coding: utf-8 -*-
"""
药品供应商价格比较分析工具
分析需要进货的药品应该从哪个供应商进货价格最低
"""

import re
from openpyxl import load_workbook
from typing import Optional, List, Dict, Tuple
from collections import defaultdict

# 文件路径
ORDER_FILE = "1.25订货.xlsx"  # 订货清单
SUPPLIER1_FILE = "瑞康志德.xlsx"  # 供应商1：瑞康志德
SUPPLIER2_FILE = "润彩.xlsx"  # 供应商2：润彩

# 列名映射
# 订货文件：品名、计划量
# 瑞康志德：商品通用名、(配)最低限价、规格
# 润彩：通用名称、最低限价、商品规格

def normalize_name(name: Optional[str]) -> str:
    """规范化药品名称，用于匹配"""
    if name is None:
        return ""
    name = str(name).strip()
    # 移除括号及括号内容
    name = re.sub(r'[\(\[（【].*?[\)\]）】]', '', name)
    # 移除多余空格
    name = re.sub(r'\s+', '', name)
    # 转小写
    name = name.lower()
    return name

def extract_number_from_spec(spec: Optional[str]) -> Tuple[float, str]:
    """
    从规格字符串中提取数字和单位
    返回: (数量, 单位)
    例如: "10片" -> (10.0, "片"), "100ml" -> (100.0, "ml")
    """
    if spec is None:
        return (1.0, "")
    
    spec = str(spec).strip()
    
    # 常见单位
    units = ['片', '粒', '丸', '袋', '盒', '瓶', '支', 'ml', 'g', 'mg', 'kg', 'l', '包', '条', '个']
    
    # 尝试提取数字
    # 匹配开头的数字（可能包含小数点）
    match = re.match(r'^(\d+\.?\d*)', spec)
    if match:
        num = float(match.group(1))
        # 提取单位
        remaining = spec[match.end():].strip()
        unit = ""
        for u in units:
            if remaining.startswith(u):
                unit = u
                break
        return (num, unit)
    
    # 如果没有找到数字，返回默认值
    return (1.0, "")

def parse_price(price) -> Optional[float]:
    """解析价格，转换为浮点数"""
    if price is None:
        return None
    try:
        if isinstance(price, (int, float)):
            return float(price)
        price_str = str(price).strip()
        # 移除可能的货币符号和空格
        price_str = price_str.replace('¥', '').replace('￥', '').replace(',', '').strip()
        return float(price_str)
    except (ValueError, TypeError):
        return None

def find_column_index(sheet, keywords: List[str]) -> Optional[int]:
    """在表头行中查找包含指定关键词的列索引（从1开始）"""
    for col_idx, cell in enumerate(sheet[1], start=1):
        if cell.value:
            cell_str = str(cell.value)
            for keyword in keywords:
                if keyword in cell_str:
                    return col_idx
    return None

def load_order_file(file_path: str) -> List[Dict]:
    """加载订货文件"""
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    
    # 查找列
    name_col = find_column_index(ws, ["品名"])
    qty_col = find_column_index(ws, ["计划量"])
    
    if not name_col or not qty_col:
        raise ValueError(f"无法找到必要的列：品名={name_col}, 计划量={qty_col}")
    
    orders = []
    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row=row_idx, column=name_col).value
        qty = ws.cell(row=row_idx, column=qty_col).value
        
        if name and qty:
            try:
                qty = float(qty) if isinstance(qty, (int, float)) else None
            except:
                qty = None
            
            if qty is not None:
                orders.append({
                    'name': str(name).strip(),
                    'normalized_name': normalize_name(name),
                    'quantity': qty,
                    'row': row_idx
                })
    
    wb.close()
    return orders

def load_supplier_file(file_path: str, supplier_name: str) -> Dict[str, List[Dict]]:
    """
    加载供应商文件
    返回: {normalized_name: [产品信息列表]}
    """
    wb = load_workbook(file_path, data_only=True)
    ws = wb.active
    
    # 根据供应商名称确定列名
    if supplier_name == "瑞康志德":
        name_keywords = ["商品通用名"]
        price_keywords = ["(配)最低限价", "最低限价"]
        spec_keywords = ["规格"]
    else:  # 润彩
        name_keywords = ["通用名称"]
        price_keywords = ["最低限价"]
        spec_keywords = ["商品规格"]
    
    name_col = find_column_index(ws, name_keywords)
    price_col = find_column_index(ws, price_keywords)
    spec_col = find_column_index(ws, spec_keywords)
    
    if not name_col:
        raise ValueError(f"在{supplier_name}文件中无法找到品名列")
    if not price_col:
        raise ValueError(f"在{supplier_name}文件中无法找到价格列")
    if not spec_col:
        raise ValueError(f"在{supplier_name}文件中无法找到规格列")
    
    products = defaultdict(list)
    
    for row_idx in range(2, ws.max_row + 1):
        name = ws.cell(row=row_idx, column=name_col).value
        price = ws.cell(row=row_idx, column=price_col).value
        spec = ws.cell(row=row_idx, column=spec_col).value
        
        if name:
            price_float = parse_price(price)
            if price_float is not None:
                spec_num, spec_unit = extract_number_from_spec(spec)
                normalized = normalize_name(name)
                
                products[normalized].append({
                    'raw_name': str(name).strip(),
                    'price': price_float,
                    'spec': str(spec).strip() if spec else "",
                    'spec_num': spec_num,
                    'spec_unit': spec_unit,
                    'unit_price': price_float / spec_num if spec_num > 0 else price_float,
                    'row': row_idx
                })
    
    wb.close()
    return products

def find_best_match(order_name: str, order_normalized: str, supplier_products: Dict[str, List[Dict]]) -> Optional[Dict]:
    """
    在供应商产品中查找最佳匹配
    返回单位价格最低的产品信息
    """
    # 精确匹配
    if order_normalized in supplier_products:
        products = supplier_products[order_normalized]
        if products:
            # 返回单位价格最低的产品
            best = min(products, key=lambda x: x['unit_price'])
            return best
    
    # 模糊匹配：检查是否包含关键词
    order_keywords = set(order_normalized)
    best_match = None
    best_score = 0
    
    for normalized, products in supplier_products.items():
        # 计算相似度（简单的关键词重叠度）
        common_chars = len(order_keywords & set(normalized))
        if common_chars > best_score and common_chars >= len(order_keywords) * 0.6:
            best_score = common_chars
            best_product = min(products, key=lambda x: x['unit_price'])
            best_match = best_product
    
    return best_match

def analyze_suppliers():
    """主分析函数"""
    print("=" * 80)
    print("药品供应商价格比较分析")
    print("=" * 80)
    print()
    
    # 加载数据
    print("正在加载订货文件...")
    orders = load_order_file(ORDER_FILE)
    print(f"  找到 {len(orders)} 个需要进货的药品")
    print()
    
    print("正在加载供应商1（瑞康志德）文件...")
    supplier1_products = load_supplier_file(SUPPLIER1_FILE, "瑞康志德")
    print(f"  找到 {len(supplier1_products)} 个不同的药品")
    print()
    
    print("正在加载供应商2（润彩）文件...")
    supplier2_products = load_supplier_file(SUPPLIER2_FILE, "润彩")
    print(f"  找到 {len(supplier2_products)} 个不同的药品")
    print()
    
    # 分析每个订单
    results = []
    not_found = []
    
    print("=" * 80)
    print("开始分析每个药品的最优供应商...")
    print("=" * 80)
    print()
    
    for order in orders:
        order_name = order['name']
        order_normalized = order['normalized_name']
        order_qty = order['quantity']
        
        # 在两个供应商中查找匹配
        match1 = find_best_match(order_name, order_normalized, supplier1_products)
        match2 = find_best_match(order_name, order_normalized, supplier2_products)
        
        print(f"药品: {order_name} (计划量: {order_qty})")
        print(f"  规范化名称: {order_normalized}")
        
        if match1:
            print(f"  瑞康志德: 找到匹配")
            print(f"    - 品名: {match1['raw_name']}")
            print(f"    - 规格: {match1['spec']}")
            print(f"    - 价格: {match1['price']:.2f}元")
            print(f"    - 单位价格: {match1['unit_price']:.4f}元/{match1['spec_unit'] if match1['spec_unit'] else '单位'}")
        else:
            print(f"  瑞康志德: 未找到匹配")
        
        if match2:
            print(f"  润彩: 找到匹配")
            print(f"    - 品名: {match2['raw_name']}")
            print(f"    - 规格: {match2['spec']}")
            print(f"    - 价格: {match2['price']:.2f}元")
            print(f"    - 单位价格: {match2['unit_price']:.4f}元/{match2['spec_unit'] if match2['spec_unit'] else '单位'}")
        else:
            print(f"  润彩: 未找到匹配")
        
        # 确定最优供应商
        if match1 and match2:
            if match1['unit_price'] < match2['unit_price']:
                best_supplier = "瑞康志德"
                best_match = match1
                savings = (match2['unit_price'] - match1['unit_price']) * order_qty
            elif match2['unit_price'] < match1['unit_price']:
                best_supplier = "润彩"
                best_match = match2
                savings = (match1['unit_price'] - match2['unit_price']) * order_qty
            else:
                best_supplier = "价格相同，任选其一"
                best_match = match1
                savings = 0
        elif match1:
            best_supplier = "瑞康志德"
            best_match = match1
            savings = None
        elif match2:
            best_supplier = "润彩"
            best_match = match2
            savings = None
        else:
            best_supplier = None
            best_match = None
            savings = None
            not_found.append(order)
        
        if best_supplier:
            print(f"  ✓ 推荐供应商: {best_supplier}")
            if savings is not None and savings > 0:
                print(f"  ✓ 预计节省: {savings:.2f}元")
            total_cost = best_match['unit_price'] * order_qty
            print(f"  ✓ 预计总成本: {total_cost:.2f}元")
        else:
            print(f"  ✗ 两个供应商都未找到此药品")
            not_found.append(order)
        
        results.append({
            'order': order,
            'match1': match1,
            'match2': match2,
            'best_supplier': best_supplier,
            'best_match': best_match,
            'savings': savings
        })
        
        print()
    
    # 汇总
    print("=" * 80)
    print("分析结果汇总")
    print("=" * 80)
    print()
    
    supplier1_count = sum(1 for r in results if r['best_supplier'] == "瑞康志德")
    supplier2_count = sum(1 for r in results if r['best_supplier'] == "润彩")
    same_price_count = sum(1 for r in results if r['best_supplier'] and "价格相同" in r['best_supplier'])
    
    print(f"推荐从瑞康志德进货: {supplier1_count} 个药品")
    print(f"推荐从润彩进货: {supplier2_count} 个药品")
    if same_price_count > 0:
        print(f"价格相同（任选）: {same_price_count} 个药品")
    print(f"两个供应商都未找到: {len(not_found)} 个药品")
    print()
    
    if not_found:
        print("=" * 80)
        print("未找到的药品列表（需要从其他地方查找）:")
        print("=" * 80)
        for order in not_found:
            print(f"  - {order['name']} (计划量: {order['quantity']})")
        print()
    
    # 详细推荐列表
    print("=" * 80)
    print("详细推荐列表")
    print("=" * 80)
    print()
    print(f"{'药品名称':<30} {'计划量':<10} {'推荐供应商':<15} {'规格':<20} {'单价':<15} {'总成本':<15}")
    print("-" * 110)
    
    for r in results:
        if r['best_supplier']:
            order = r['order']
            match = r['best_match']
            total_cost = match['unit_price'] * order['quantity']
            print(f"{order['name']:<30} {order['quantity']:<10} {r['best_supplier']:<15} {match['spec']:<20} {match['unit_price']:.4f}元/{match['spec_unit'] if match['spec_unit'] else '单位':<10} {total_cost:.2f}元")
    
    if not_found:
        print()
        print("未找到的药品:")
        for order in not_found:
            print(f"{order['name']:<30} {order['quantity']:<10} {'未找到':<15}")

if __name__ == "__main__":
    try:
        analyze_suppliers()
    except Exception as e:
        import traceback
        print(f"发生错误: {e}")
        traceback.print_exc()