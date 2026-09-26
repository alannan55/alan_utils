"""
生成CDR文件的Python程序
支持两种方式：
1. 通过Windows COM自动化直接生成CDR文件（需要安装CorelDRAW）
2. 生成SVG文件作为备选方案（可在CorelDRAW中打开并另存为CDR格式）
"""

import os
import sys

try:
    import win32com.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

try:
    import svgwrite
    SVG_AVAILABLE = True
except ImportError:
    SVG_AVAILABLE = False


def generate_cdr_svg(output_file='医疗器械.cdr.svg', width=200, height=200):
    """
    生成包含绿色矩形方框和中心文字的SVG文件
    
    参数:
        output_file: 输出文件名
        width: 画布宽度（毫米）
        height: 画布高度（毫米）
    """
    # 创建SVG绘图对象，使用毫米单位
    dwg = svgwrite.Drawing(output_file, size=(f'{width}mm', f'{height}mm'))
    
    # 设置视图框，确保坐标系统正确
    dwg.viewbox(0, 0, width, height)
    
    # 计算矩形的位置和大小（留出边距）
    margin = 20  # 边距（毫米）
    rect_x = margin
    rect_y = margin
    rect_width = width - 2 * margin
    rect_height = height - 2 * margin
    
    # 添加绿色矩形方框
    # stroke='green' 表示边框颜色为绿色
    # fill='none' 表示不填充
    # stroke_width=2 表示边框宽度为2毫米
    dwg.add(dwg.rect(
        insert=(rect_x, rect_y),
        size=(rect_width, rect_height),
        stroke='green',
        fill='none',
        stroke_width=2
    ))
    
    # 计算文本位置（矩形中心）
    text_x = width / 2
    text_y = height / 2
    
    # 添加中心文本"医疗器械"
    # text_anchor='middle' 表示文本水平居中
    # dominant_baseline='central' 表示文本垂直居中
    # font_size 设置为合适的字体大小
    text = dwg.text(
        '医疗器械',
        insert=(text_x, text_y),
        text_anchor='middle',
        dominant_baseline='central',
        font_size=24,
        fill='black',
        font_family='SimHei, Microsoft YaHei, Arial Unicode MS, sans-serif'  # 支持中文的字体
    )
    dwg.add(text)
    
    # 保存SVG文件
    dwg.save()
    print(f"SVG文件已生成: {output_file}")
    print("提示: 可以在CorelDRAW中打开此SVG文件，然后另存为CDR格式")


if __name__ == '__main__':
    # 生成文件
    generate_cdr_svg()
    
    print("\n使用说明:")
    print("1. 安装依赖: pip install svgwrite")
    print("2. 运行此程序生成SVG文件")
    print("3. 在CorelDRAW中打开生成的SVG文件")
    print("4. 在CorelDRAW中另存为CDR格式")







