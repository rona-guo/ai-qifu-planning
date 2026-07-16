"""Generate architecture diagram PNG for Tiansui proposal"""
from PIL import Image, ImageDraw, ImageFont
import os

WIDTH, HEIGHT = 1200, 920
BG = (255, 255, 255)

# Colors
C_BLUE = {"fill": (230, 241, 251), "stroke": (24, 95, 165), "text": (12, 68, 124)}
C_PURPLE = {"fill": (238, 237, 254), "stroke": (83, 74, 183), "text": (60, 52, 137)}
C_TEAL = {"fill": (225, 245, 238), "stroke": (15, 110, 86), "text": (8, 80, 65)}
C_AMBER = {"fill": (250, 238, 218), "stroke": (133, 79, 11), "text": (99, 56, 6)}
C_GRAY = {"fill": (241, 239, 232), "stroke": (95, 94, 90), "text": (68, 68, 65)}
C_WHITE = (255, 255, 255)

# Fonts
try:
    font_title = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttc", 26)
    font_label = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttc", 18)
    font_body = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 15)
    font_small = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 13)
    font_badge = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttc", 13)
except:
    font_title = ImageFont.load_default()
    font_label = ImageFont.load_default()
    font_body = ImageFont.load_default()
    font_small = ImageFont.load_default()
    font_badge = ImageFont.load_default()

img = Image.new("RGB", (WIDTH, HEIGHT), BG)
draw = ImageDraw.Draw(img)

MARGIN = 60
LAYER_W = WIDTH - 2 * MARGIN

# Title
draw.text((WIDTH // 2, 30), "系统技术架构", fill=(38, 33, 92), font=font_title, anchor="mm")

def draw_layer(y, h, colors, label, content_text, sub_content=None):
    """Draw a layer rectangle with label and content"""
    x = MARGIN
    # Background
    draw.rounded_rectangle([x, y, x + LAYER_W, y + h], radius=10,
                           fill=colors["fill"], outline=colors["stroke"], width=2)
    # Label
    draw.text((x + 20, y + h // 2), label, fill=colors["text"], font=font_label, anchor="lm")
    # Content
    if content_text:
        cx = x + 130
        draw.text((cx + (LAYER_W - 130) // 2, y + h // 2), content_text,
                  fill=colors["text"], font=font_body, anchor="mm")
    if sub_content:
        draw.text((x + 130 + (LAYER_W - 130) // 2, y + h - 15), sub_content,
                  fill=colors["stroke"], font=font_small, anchor="mm")

def draw_module_box(x, y, w, h, text, colors):
    """Draw a module box inside a layer"""
    draw.rounded_rectangle([x, y, x + w, y + h], radius=6,
                           fill=colors["stroke"], outline=colors["stroke"], width=1)
    # Lighter fill
    light_fill = tuple(min(255, c + 30) for c in colors["stroke"])
    draw.rounded_rectangle([x + 1, y + 1, x + w - 1, y + h - 1], radius=6,
                           fill=(206, 203, 246), outline=colors["stroke"], width=1)
    draw.text((x + w // 2, y + h // 2), text, fill=colors["text"],
              font=font_small, anchor="mm")

def draw_badge(x, y, w, h, text, fill_color, text_color):
    """Draw a data classification badge"""
    draw.rounded_rectangle([x, y, x + w, y + h], radius=4,
                           fill=fill_color, outline=fill_color, width=1)
    draw.text((x + w // 2, y + h // 2), text, fill=text_color,
              font=font_badge, anchor="mm")

# Layer 1: Access
y = 70
draw_layer(y, 55, C_BLUE, "访问层", "企微 / PC / 移动端  —  统一入口，单点登录")

# Layer 2: Application (taller, with module boxes)
y = 140
h = 130
x = MARGIN
draw.rounded_rectangle([x, y, x + LAYER_W, y + h], radius=10,
                       fill=C_PURPLE["fill"], outline=C_PURPLE["stroke"], width=2)
draw.text((x + 20, y + 25), "应用层", fill=C_PURPLE["text"], font=font_label, anchor="lm")
draw.text((x + 100, y + 25), "明道云", fill=C_PURPLE["stroke"], font=font_body, anchor="lm")

# Module boxes
modules = ["基础管理", "项目管理", "报告审批", "知识库", "培训管理"]
box_w = 120
box_h = 40
box_y = y + 55
gap = (LAYER_W - 140 - 5 * box_w) // 4
start_x = x + 70
for i, mod in enumerate(modules):
    bx = start_x + i * (box_w + gap)
    draw.rounded_rectangle([bx, box_y, bx + box_w, box_y + box_h], radius=6,
                           fill=(206, 203, 246), outline=C_PURPLE["stroke"], width=1)
    draw.text((bx + box_w // 2, box_y + box_h // 2), mod,
              fill=C_PURPLE["text"], font=font_small, anchor="mm")

draw.text((x + LAYER_W // 2, y + h - 15), "工作流引擎",
          fill=C_PURPLE["stroke"], font=font_small, anchor="mm")

# Layer 3: AI
y = 285
draw_layer(y, 55, C_TEAL, "AI能力层", "明道云自带AI：模板填充 / 工作流自动化 / AI助手 / 知识检索")

# Layer 4: Data (taller, with badges)
y = 355
h = 90
x = MARGIN
draw.rounded_rectangle([x, y, x + LAYER_W, y + h], radius=10,
                       fill=C_AMBER["fill"], outline=C_AMBER["stroke"], width=2)
draw.text((x + 20, y + 25), "数据层", fill=C_AMBER["text"], font=font_label, anchor="lm")

# Data types
data_types = ["业务数据", "文件附件", "知识库"]
dt_y = y + 25
dt_x = x + 130
for dt in data_types:
    draw.text((dt_x, dt_y), dt, fill=C_AMBER["stroke"], font=font_body, anchor="lm")
    dt_x += 130

# L1-L4 badges
badge_y = y + 55
badge_w = 110
badge_h = 24
badges = [
    ("L1 公开", (250, 199, 117), (99, 56, 6)),
    ("L2 内部", (250, 199, 117), (99, 56, 6)),
    ("L3 敏感", (239, 159, 39), C_WHITE),
    ("L4 核心", (133, 79, 11), C_WHITE),
]
bx = x + 130
gap = 15
for text, fill_c, text_c in badges:
    draw_badge(bx, badge_y, badge_w, badge_h, text, fill_c, text_c)
    bx += badge_w + gap

# Layer 5: Security
y = 460
draw_layer(y, 55, C_GRAY, "安全层", "私有化部署  |  权限分级  |  操作留痕  |  数据加密")

# Logic loop at bottom
y = 540
draw.text((WIDTH // 2, y), "数据逻辑闭环", fill=(38, 33, 92), font=font_label, anchor="mm")

loop_steps = [
    "客户文件", "AI提取", "人工核准", "入库", "报告生成",
    "AI审核", "审批", "归档", "知识库积累", "复用"
]
step_y = y + 40
step_w = 100
total_w = len(loop_steps) * step_w
start_x = (WIDTH - total_w) // 2
for i, step in enumerate(loop_steps):
    sx = start_x + i * step_w
    # Box
    box_color = C_BLUE if i < 4 else (C_PURPLE if i < 7 else C_TEAL)
    draw.rounded_rectangle([sx + 5, step_y, sx + step_w - 5, step_y + 36],
                           radius=6, fill=box_color["fill"], outline=box_color["stroke"], width=1)
    draw.text((sx + step_w // 2, step_y + 18), step,
              fill=box_color["text"], font=font_small, anchor="mm")
    # Arrow
    if i < len(loop_steps) - 1:
        ax = sx + step_w - 5
        ay = step_y + 18
        draw.line([(ax, ay), (ax + 10, ay)], fill=(136, 135, 128), width=2)
        # Arrow head
        draw.polygon([(ax + 10, ay - 4), (ax + 14, ay), (ax + 10, ay + 4)],
                     fill=(136, 135, 128))

# Loop back arrow
back_y = step_y + 55
draw.line([(start_x + step_w // 2, step_y + 36), (start_x + step_w // 2, back_y),
           (start_x + (len(loop_steps) - 1) * step_w + step_w // 2, back_y),
           (start_x + (len(loop_steps) - 1) * step_w + step_w // 2, step_y + 36)],
          fill=(136, 135, 128), width=1)
draw.text((WIDTH // 2, back_y + 12), "经验复用 → 下次同类项目自动推荐",
          fill=(136, 135, 128), font=font_small, anchor="mm")

# Save
output_path = r"D:\AI落地\5_项目交付\天遂税务师事务所AI数字化\生成产物\技术架构图.png"
img.save(output_path, "PNG", dpi=(300, 300))
print(f"Saved: {output_path}")
