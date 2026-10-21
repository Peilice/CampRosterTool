"""讀取原始報名 Excel 檔（支援 .xlsx 及 SpreadsheetML 格式的 .xls）。

同一分頁內可能同時存在多組表頭不同的資料區塊（網路報名系統匯出的資料 +
Google 表單匯出的現場報名資料），本模組會逐列掃描找出每個表頭列，
將各區塊資料統一轉換成標準欄位後再合併回傳。
"""
import re
import xml.etree.ElementTree as ET

import openpyxl

REQUIRED_HEADERS = {"學員姓名", "身分證字號", "參加梯次", "性別"}

NS = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}

CANONICAL_HEADER = [
    "來源", "項次", "學員姓名", "性別", "出生年", "月", "日", "身分證字號",
    "縣市", "郵遞區號", "地址", "家長姓名", "家長行動電話", "EMAIL",
    "學員目前就讀學校", "學員就讀年級", "參加梯次",
    "(第一梯次)是否有意願上台進行才藝表演", "以前是否參加過元亨寺暑期學佛夏令營",
    "緊急連絡人姓名", "緊急連絡人電話", "緊急連絡人與學員關係",
    "學童健康狀況（風險評估）()", "學員特殊狀況說明",
    "活動中，如有身體不適或意外，當下聯繫不上家長時：", "管理者註記事項",
]

# 現場報名（Google 表單）匯出檔才有的欄位，用來區分來源（功能規格 §3.4）
ONSITE_MARKER_HEADER = "時間戳記"

# 部分報名表單（如 Google 表單匯出）欄位名稱與標準欄位不同，需要對應轉換。
# key 為去除開頭數字編號、前後空白後的欄位文字；value 為對應的標準欄位名稱。
HEADER_ALIASES = {
    "生日：(民國年)/月/日（例：104/05/06）": "出生年",
    "身份證字號": "身分證字號",
    "通訊地址": "地址",
    "學員目前就讀學校(縣市/校名)": "學員目前就讀學校",
    "報名梯次": "參加梯次",
    "緊急聯絡人姓名": "緊急連絡人姓名",
    "健康狀況（風險評估）": "學童健康狀況（風險評估）()",
    "活動中有身體不適或意外，當下聯繫不上家長時：": "活動中，如有身體不適或意外，當下聯繫不上家長時：",
}

_LEADING_NUM_RE = re.compile(r"^\d+[.．]\s*")

NO_HEADER_MESSAGE = (
    "在檔案中找不到報名資料。請確認上傳的是報名系統匯出的原始檔，"
    "且資料表包含「學員姓名」、「身分證字號」、「參加梯次」、「性別」這幾欄。"
)


class SourceFormatError(ValueError):
    """原始檔格式不符，訊息可直接顯示給使用者。"""


def _normalize_header_cell(text):
    text = (text or "").strip()
    text = _LEADING_NUM_RE.sub("", text)
    return HEADER_ALIASES.get(text, text)


def _extract_blocks(all_rows, row_nos):
    """掃描整份工作表（list[list[str]]），找出所有表頭列並合併各區塊資料。

    row_nos 為 all_rows 各列在工作表中的列號（與 Excel 左側列號相同）。

    表頭列以「去除數字編號、套用欄位別名後，是否涵蓋 REQUIRED_HEADERS」判斷。
    來源依表頭格式決定：有「時間戳記」欄的是現場報名，其餘是網路報名。
    找不到對應的欄位（如現場報名沒有家長姓名、EMAIL）則留空。
    """
    header_positions = []
    for i, row in enumerate(all_rows):
        normalized = [_normalize_header_cell(c) for c in row]
        if REQUIRED_HEADERS.issubset(set(normalized)):
            header_positions.append((i, normalized))
    if not header_positions:
        return None

    rows = []
    out_row_nos = []
    for block_idx, (start, normalized_header) in enumerate(header_positions):
        end = (
            header_positions[block_idx + 1][0]
            if block_idx + 1 < len(header_positions)
            else len(all_rows)
        )
        source = "現場報名" if ONSITE_MARKER_HEADER in normalized_header else "網路報名"
        for row, row_no in zip(all_rows[start + 1:end], row_nos[start + 1:end]):
            if all((v or "").strip() == "" for v in row):
                continue
            raw = dict(zip(normalized_header, row))
            raw["來源"] = source
            rows.append([raw.get(col, "") for col in CANONICAL_HEADER])
            out_row_nos.append(row_no)
    return CANONICAL_HEADER, rows, out_row_nos


def _rows_from_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        all_rows = [
            ["" if v is None else str(v) for v in row]
            for row in ws.iter_rows(values_only=True)
        ]
        result = _extract_blocks(all_rows, list(range(1, len(all_rows) + 1)))
        if result is not None:
            return result
    raise SourceFormatError(NO_HEADER_MESSAGE)


def _rows_from_spreadsheetml(path):
    with open(path, encoding="utf-8") as f:
        content = f.read()
    root = ET.fromstring(content)
    for ws in root.findall("ss:Worksheet", NS):
        table = ws.find("ss:Table", NS)
        if table is None:
            continue
        rows_xml = table.findall("ss:Row", NS)
        if not rows_xml:
            continue

        def read_row(row_xml):
            values = []
            next_index = 1
            for cell in row_xml.findall("ss:Cell", NS):
                idx_attr = cell.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
                if idx_attr is not None:
                    idx = int(idx_attr)
                    while next_index < idx:
                        values.append("")
                        next_index += 1
                data = cell.find("ss:Data", NS)
                values.append(data.text if data is not None and data.text is not None else "")
                next_index += 1
            return values

        # <Row> 可用 ss:Index 跳過空白列，列號需依此累計
        all_rows = []
        row_nos = []
        next_row_no = 1
        for row_xml in rows_xml:
            idx_attr = row_xml.get("{urn:schemas-microsoft-com:office:spreadsheet}Index")
            if idx_attr is not None:
                next_row_no = int(idx_attr)
            all_rows.append(read_row(row_xml))
            row_nos.append(next_row_no)
            next_row_no += 1
        result = _extract_blocks(all_rows, row_nos)
        if result is not None:
            return result
    raise SourceFormatError(NO_HEADER_MESSAGE)


def read_source(path):
    """回傳 (header: list[str], rows: list[list[str]], row_nos: list[int])"""
    if path.lower().endswith(".xlsx"):
        return _rows_from_xlsx(path)

    with open(path, "rb") as f:
        head = f.read(200)
    if b"<?xml" in head:
        return _rows_from_spreadsheetml(path)
    return _rows_from_xlsx(path)
