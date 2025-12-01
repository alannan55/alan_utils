#!/usr/bin/env python3
"""
png_to_svg.py

将 PNG 位图转换为 SVG 矢量图的简单脚本。
方法：
  - 使用 OpenCV 找到二值/边缘轮廓（支持透明通道优先）
  - 使用 svgwrite 将轮廓转换为 SVG 路径，并计算每个轮廓内的平均颜色作为填充色

依赖：
  pip install pillow numpy opencv-python svgwrite

注意：此脚本进行位图轮廓矢量化（基于轮廓的多边形逼近），并非像手工矢量绘制那样保持完美平滑度。
对于复杂或照片级图像，结果会比较粗糙；更好的工具包括 Potrace、Inkscape 的自动追踪功能或专用 raster-to-vector 软件。

用法：在 __main__ 中通过变量赋值指定输入/输出路径及参数。
"""

from typing import Tuple, Optional
import numpy as np
from PIL import Image
import cv2
import svgwrite
import math


def _rgba_to_tuple(color: np.ndarray) -> Tuple[int, int, int, float]:
    """将 RGBA 数组转换为 (r,g,b,a) 格式，alpha 为 0..1。"""
    r, g, b, a = color
    return int(r), int(g), int(b), float(a) / 255.0


def _mean_color_in_mask(rgba: np.ndarray, mask: np.ndarray) -> Tuple[int,int,int,float]:
    """计算 rgba 图像在布尔 mask 区域内的平均颜色（忽略透明像素）。
    mask: 单通道布尔数组，形状 (h,w)，表示当前轮廓内的像素。
    返回 (r,g,b,a) 其中 a 在 0..1
    """
    if mask.sum() == 0:
        return (0,0,0,0.0)
    # 取出对应像素
    pixels = rgba[mask]
    # 计算平均 RGBA
    mean = pixels.mean(axis=0)
    r,g,b,a = mean
    return int(r), int(g), int(b), float(a) / 255.0


def _rgba_to_svg_fill(rgba_tuple: Tuple[int,int,int,float]) -> Optional[str]:
    r,g,b,a = rgba_tuple
    if a <= 0.001:
        return None
    if a >= 0.999:
        return f'rgb({r},{g},{b})'
    # 使用 rgba() CSS 格式
    return f'rgba({r},{g},{b},{a:.3f})'


def contours_to_svg_paths(contours, hierarchy, rgba_img: np.ndarray, simplify_tolerance: float = 2.0, min_area: int = 20):
    """
    将 OpenCV 轮廓转换为 SVG 路径数据和样式（填充颜色）
    参数:
      contours: OpenCV 找到的轮廓列表
      hierarchy: OpenCV 层次结构（可为 None）
      rgba_img: 原始图像的 numpy 数组 (H,W,4)
      simplify_tolerance: 多边形逼近的 epsilon 参数（像素）
      min_area: 忽略小于此面积的轮廓
    返回: 列表 of dict { 'd': path_d, 'fill': fill_css, 'stroke': stroke_css }
    """
    paths = []
    h, w = rgba_img.shape[:2]
    for i, cnt in enumerate(contours):
        area = cv2.contourArea(cnt)
        if area < min_area:
            continue
        # 多边形逼近，epsilon 和周长相关
        peri = cv2.arcLength(cnt, True)
        epsilon = max(1.0, simplify_tolerance * (peri / 100.0))
        approx = cv2.approxPolyDP(cnt, epsilon, True)
        # 构造 mask 用来计算平均颜色
        mask = np.zeros((h, w), dtype=np.uint8)
        cv2.drawContours(mask, [approx], -1, color=255, thickness=-1)
        mask_bool = mask.astype(bool)
        mean_rgba = _mean_color_in_mask(rgba_img, mask_bool)
        fill_css = _rgba_to_svg_fill(mean_rgba)
        # 构造 path d 字符串
        pts = approx.reshape(-1, 2)
        if pts.shape[0] < 2:
            continue
        d_parts = [f'M {pts[0][0]:.2f},{pts[0][1]:.2f}']
        for (x,y) in pts[1:]:
            d_parts.append(f'L {x:.2f},{y:.2f}')
        d_parts.append('Z')
        d = ' '.join(d_parts)
        # 如果没有有效填充色，将使用描边（使用平均非透明像素颜色）
        stroke_css = None
        if fill_css is None:
            # 找到轮廓上像素的平均颜色作为描边颜色
            # 生成轮廓边掩码
            edge_mask = np.zeros((h, w), dtype=np.uint8)
            cv2.drawContours(edge_mask, [approx], -1, color=255, thickness=1)
            edge_mask_bool = edge_mask.astype(bool)
            mean_edge = _mean_color_in_mask(rgba_img, edge_mask_bool)
            stroke_css = _rgba_to_svg_fill(mean_edge)
        paths.append({'d': d, 'fill': fill_css, 'stroke': stroke_css, 'area': area})
    # 按面积从大到小排序，方便先绘制大的
    paths.sort(key=lambda x: -x['area'])
    return paths


def png_to_svg(input_png: str, output_svg: str, simplify_tolerance: float = 2.0, min_area: int = 20, background: Optional[Tuple[int,int,int]] = None, svg_size: Optional[Tuple[int,int]] = None):
    """
    将 PNG 转为 SVG（基于轮廓逼近）

    参数:
      input_png: 输入 PNG 路径
      output_svg: 输出 SVG 路径
      simplify_tolerance: 多边形简化的容忍度（越大越简化）
      min_area: 忽略小面积轮廓
      background: 如果提供 (r,g,b) 则在 SVG 背景上绘制该颜色矩形；若为 None，则保持透明背景
      svg_size: 可选的 (width, height) 覆盖输出的 SVG 画布尺寸（像素）
    """
    img = Image.open(input_png).convert('RGBA')
    rgba = np.array(img)
    h, w = rgba.shape[:2]

    # 使用 alpha 通道优先，创建二值图像用于轮廓检测
    alpha = rgba[:, :, 3]
    # 如果存在透明度，优先用 alpha 阈值
    if alpha.max() < 255:
        # 有透明区域——把非透明作为前景
        bw = (alpha > 10).astype('uint8') * 255
    else:
        # 没有透明度，转换为灰度并阈值
        gray = cv2.cvtColor(rgba[:, :, :3], cv2.COLOR_RGB2GRAY)
        # 自适应阈值可以在不同亮度下更稳健
        bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 15, -10)
    # 可选：对 bw 做形态学操作，去噪
    kernel = np.ones((3,3), np.uint8)
    bw = cv2.morphologyEx(bw, cv2.MORPH_OPEN, kernel)

    # 查找轮廓
    contours_info = cv2.findContours(bw, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = contours_info[0] if len(contours_info) == 2 else contours_info[1]

    paths = contours_to_svg_paths(contours, None, rgba, simplify_tolerance=simplify_tolerance, min_area=min_area)

    svg_w, svg_h = (w, h) if svg_size is None else svg_size
    dwg = svgwrite.Drawing(output_svg, size=(f'{svg_w}px', f'{svg_h}px'))

    if background is not None:
        r,g,b = background
        dwg.add(dwg.rect(insert=(0,0), size=(svg_w, svg_h), fill=f'rgb({r},{g},{b})'))

    # 将路径逐一加入
    for p in paths:
        fill = p['fill']
        stroke = p['stroke']
        if fill is not None:
            dwg.add(dwg.path(d=p['d'], fill=fill, stroke='none'))
        else:
            # 没有填充色则画描边
            if stroke is None:
                stroke = 'black'
            dwg.add(dwg.path(d=p['d'], fill='none', stroke=stroke, stroke_width=1))

    # 保存文件
    dwg.save()


if __name__ == '__main__':
    # ---------- 用户可以在这里修改输入参数 ----------
    INPUT_PNG = r"E:\项目\诚善堂\logo\诚善堂LOGO字图.png"        # 输入 PNG 路径
    OUTPUT_SVG = r"E:\项目\诚善堂\logo\诚善堂LOGO字图.svg"      # 输出 SVG 路径

    # 多边形简化容忍度（越大输出越简化）
    SIMPLIFY_TOLERANCE = 2.0
    # 忽略小面积（像素）
    MIN_AREA = 40
    # 如果想要白色背景可以设为 (255,255,255)，否则设为 None 保留透明背景
    BACKGROUND = None
    # 如果想输出不同尺寸，可以给一个 (width, height)，否则使用 PNG 原尺寸
    SVG_SIZE = None
    # ------------------------------------------------

    print(f'Converting "{INPUT_PNG}" -> "{OUTPUT_SVG}"')
    png_to_svg(INPUT_PNG, OUTPUT_SVG, simplify_tolerance=SIMPLIFY_TOLERANCE, min_area=MIN_AREA, background=BACKGROUND, svg_size=SVG_SIZE)
    print('Done.')
