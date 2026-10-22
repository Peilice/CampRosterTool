"""類型 E：公告.png（公開張貼的錄取名單圖片）"""
import os

from PIL import Image, ImageDraw, ImageFont

from processing import SESSION_LABEL, group_students, mask_id, mask_name

# 網頁版由 index.html 下載字型寫到這裡；本機測試可改成系統字型路徑。
FONT_PATH = "/app/NotoSansTC-Regular.otf"

ROWS_PER_BLOCK = 50
HEADERS = ("編號", "姓名", "身分證字號")

FONT_SIZE = 26
TITLE_SIZE = 44
ROW_HEIGHT = 42
CELL_PADDING = 18
BLOCK_GAP = 36
MARGIN = 40
TITLE_HEIGHT = 90

BORDER_COLOR = "#000000"
HEADER_FILL = "#1F3864"
HEADER_TEXT = "#FFFFFF"

GRID_FILL_COLOR = {
    (1, "男"): "#CCFFCC",  # 淺綠
    (1, "女"): "#FFFFCC",  # 淺黃
    (2, "男"): "#CCECFF",  # 淺藍
    (2, "女"): "#FFCCCC",  # 淺紅
}


def _col_widths(font, rows):
    return [
        int(max(font.getlength(str(v)) for v in values)) + 2 * CELL_PADDING
        for values in zip(HEADERS, *rows)
    ]


def _draw_row(draw, font, x, y, widths, values, fill, text_color):
    for width, value in zip(widths, values):
        draw.rectangle((x, y, x + width, y + ROW_HEIGHT), fill=fill, outline=BORDER_COLOR, width=2)
        draw.text((x + width / 2, y + ROW_HEIGHT / 2), str(value), font=font, fill=text_color, anchor="mm")
        x += width


def render(title, students, fill_color):
    """畫出一個組別的公告圖：每 50 人一個區塊，由左到右排列，姓名與身分證字號遮蔽。"""
    font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
    title_font = ImageFont.truetype(FONT_PATH, TITLE_SIZE)
    rows = [(s.code, mask_name(s.get("學員姓名")), mask_id(s.get("身分證字號"))) for s in students]
    blocks = [rows[i:i + ROWS_PER_BLOCK] for i in range(0, len(rows), ROWS_PER_BLOCK)]

    widths = _col_widths(font, rows)
    block_width = sum(widths)
    tallest = max(len(b) for b in blocks)
    content_width = len(blocks) * block_width + (len(blocks) - 1) * BLOCK_GAP
    width = max(content_width, int(title_font.getlength(title))) + 2 * MARGIN
    height = MARGIN + TITLE_HEIGHT + (tallest + 1) * ROW_HEIGHT + MARGIN

    image = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(image)
    draw.text((width / 2, MARGIN + TITLE_HEIGHT / 2), title, font=title_font, fill="#000000", anchor="mm")

    left = (width - content_width) // 2
    for b, block in enumerate(blocks):
        x = left + b * (block_width + BLOCK_GAP)
        y = MARGIN + TITLE_HEIGHT
        _draw_row(draw, font, x, y, widths, HEADERS, HEADER_FILL, HEADER_TEXT)
        for row in block:
            y += ROW_HEIGHT
            _draw_row(draw, font, x, y, widths, row, fill_color, "#000000")
    return image


def build(students_all, sessions_wanted, folder, base_name):
    """公告：依梯次、性別各自產生一張圖；沒有學員的組別不產生。"""
    groups = group_students(students_all["students"], sessions_wanted)

    output_paths = []
    for session in sorted(sessions_wanted):
        label = SESSION_LABEL[session]
        for gender in ("男", "女"):
            group = groups.get((session, gender), [])
            if not group:
                continue

            image = render(f"第{label}梯次({gender})", group, GRID_FILL_COLOR[(session, gender)])
            output_path = os.path.join(folder, f"{base_name}_第{label}梯_{gender}.png")
            image.save(output_path)
            output_paths.append(output_path)

    return output_paths
