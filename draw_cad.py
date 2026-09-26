import ezdxf
from ezdxf.enums import TextEntityAlignment

# ==================== 1. 创建DXF文件（CAD标准格式） ====================
doc = ezdxf.new(dxfversion="R2010")  # 兼容所有CAD版本
msp = doc.modelspace()  # 模型空间

# ==================== 2. 图层设置（专业规范） ====================
# 外墙（红色 0.35mm）
doc.layers.add("WALL_EXT", color=1, lineweight=35)
# 内墙（黄色 0.20mm）
doc.layers.add("WALL_INT", color=2, lineweight=20)
# 门窗（蓝色）
doc.layers.add("DOOR_WIN", color=3)
# 标注（绿色）
doc.layers.add("DIM", color=5)
# 文字（白色）
doc.layers.add("TEXT", color=7)

# ==================== 3. 户型核心参数（单位：mm） ====================
T_WALL = 240    # 承重墙厚度
T_LIGHT = 120   # 轻质墙厚度
TOTAL_W = 9600  # 总宽度
TOTAL_D = 9300  # 总深度

# ==================== 4. 绘制外墙（240厚） ====================
def draw_wall(start, end, thickness, layer):
    """绘制墙体函数"""
    x1, y1 = start
    x2, y2 = end
    dx = x2 - x1
    dy = y2 - y1
    length = (dx**2 + dy**2)**0.5
    if length == 0:
        return
    
    offset_x = -dy * thickness / length / 2
    offset_y = dx * thickness / length / 2
    
    points = [
        (x1 + offset_x, y1 + offset_y),
        (x2 + offset_x, y2 + offset_y),
        (x2 - offset_x, y2 - offset_y),
        (x1 - offset_x, y1 - offset_y),
    ]
    msp.add_lwpolyline(points, close=True, dxfattribs={"layer": layer})

# 外轮廓四面墙
draw_wall((0, 0), (TOTAL_W, 0), T_WALL, "WALL_EXT")
draw_wall((TOTAL_W, 0), (TOTAL_W, TOTAL_D), T_WALL, "WALL_EXT")
draw_wall((TOTAL_W, TOTAL_D), (0, TOTAL_D), T_WALL, "WALL_EXT")
draw_wall((0, TOTAL_D), (0, 0), T_WALL, "WALL_EXT")

# 内部承重墙
draw_wall((2800, 0), (2800, 4100), T_WALL, "WALL_EXT")
draw_wall((6400, 0), (6400, 9300), T_WALL, "WALL_EXT")
draw_wall((0, 6100), (9600, 6100), T_WALL, "WALL_EXT")
draw_wall((0, 4100), (3040, 4100), T_WALL, "WALL_EXT")
draw_wall((1400, 6100), (1400, 4340), T_WALL, "WALL_EXT")

# 内部轻质墙
draw_wall((6400, 4100), (9600, 4100), T_LIGHT, "WALL_INT")
draw_wall((1640, 4340), (3040, 5300), T_LIGHT, "WALL_INT")

# ==================== 5. 绘制门窗 ====================
def draw_door(center, width, angle, layer):
    """绘制门"""
    x, y = center
    msp.add_arc((x, y), width/2, angle, angle+90, dxfattribs={"layer": layer})
    msp.add_line((x, y), (x, y-width/2), dxfattribs={"layer": layer})

def draw_window(start, end, layer):
    """绘制窗"""
    msp.add_line(start, end, dxfattribs={"layer": layer})

# 门
draw_door((3040, 4340), 1000, 0, "DOOR_WIN")
draw_door((2800, 4340), 900, 90, "DOOR_WIN")
draw_door((9600, 6100), 900, 180, "DOOR_WIN")
draw_door((1400, 5300), 800, 270, "DOOR_WIN")
draw_door((6400, 5300), 800, 270, "DOOR_WIN")

# 南向飘窗
draw_window((0, 0), (2800, 0), "DOOR_WIN")
draw_window((3040, 0), (6640, 0), "DOOR_WIN")
draw_window((6640, 0), (9600, 0), "DOOR_WIN")

# ==================== 6. 房间文字标注 ====================
rooms = [
    ((1400, 2050), "左下卧室"),
    ((4820, 2050), "起居室"),
    ((7920, 2050), "主卧"),
    ((4820, 4700), "餐厅"),
    ((2200, 6700), "厨房"),
    ((700, 5200), "卫生间"),
    ((7920, 5200), "主卫"),
    ((7920, 7700), "卧室"),
    ((2200, 4700), "玄关"),
]

for pos, text in rooms:
    msp.add_text(
        text, height=200, dxfattribs={"layer": "TEXT"}
    ).set_placement(pos, align=TextEntityAlignment.MIDDLE_CENTER)

# ==================== 7. 指北针 ====================
msp.add_polyline2d([(8500, 9000), (8700, 9000), (8600, 9200)], close=True, dxfattribs={"layer": "TEXT"})
msp.add_text("N", height=150, dxfattribs={"layer": "TEXT"}).set_placement((8600, 8900), align=TextEntityAlignment.MIDDLE_CENTER)

# ==================== 8. 保存文件 ====================
doc.saveas("户型图_完整.dxf")
print("✅ CAD文件生成完成：户型图_完整.dxf")
