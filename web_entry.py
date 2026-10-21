"""網頁版（Pyodide）進入點：供 index.html 呼叫。

瀏覽器中沒有真實檔案系統，原始檔會先寫入 Pyodide 的虛擬檔案系統，
各 output 模組照常把檔案存到 OUT_DIR，再由網頁讀出並提供下載。
"""
import json
import os
import re
import shutil
import zipfile

import output_a
import output_b
import output_c
import output_d
from processing import (
    CANCELLED, PROMOTED, WAITLISTED, assign_codes, build_students, group_label, group_students,
)
from reader import SourceFormatError, read_source

OUT_DIR = "/out"

OUTPUT_TYPES = {
    "1": ("學員整理", "xlsx", output_a),
    "2": ("學員狀況總覽", "docx", output_b),
    "3": ("報到及查詢資料", "xlsx", output_c),
    "4": ("編隊", "xlsx", output_d),
}

SESSION_CHOICES = {
    "1": {1},
    "2": {2},
    "3": {1, 2},
}

_state = {}

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]')

# 讀取後的訊息（功能規格 §7.2、§7.3），依顯示順序排列：(issues 鍵, 標題)
BLOCKER_MESSAGES = [
    ("invalid_onsite_no",
     "有 {n} 位現場報名學員的編號不正確，無法產生報表。"
     "請在原始檔現場報名的「項次」欄填入現場發給學員的 6 碼完整編號（例如 100012），再重新上傳。"),
    ("duplicate_onsite_no",
     "有 {n} 組現場報名學員的編號重複，無法產生報表。請修正原始檔現場報名的「項次」欄後重新上傳。"),
    ("duplicate_promote_n",
     "有 {n} 組候補學員的轉正號碼重複，無法決定編號順序。"
     "請修正原始檔「管理者註記事項」的「轉正N」後重新上傳。"),
]
WARNING_MESSAGES = [
    ("unrecognized",
     "以下 {n} 筆資料缺少或無法辨識梯次、性別，未納入報表"
     "（已在「學員整理」的原檔分頁以橘色標示）。請補填後重新上傳："),
    ("conflict",
     "以下 {n} 筆現場報名資料的編號第 1 碼與實際梯次、性別不符，未納入報表"
     "（已在「學員整理」的原檔分頁以黃色標示）。請人工核對："),
    ("duplicate_id", "以下 {n} 組資料的身分證字號重複，可能是重複報名，請人工核對："),
    ("bad_birth", "以下 {n} 筆資料的出生年月日無法解讀，報表中保留原值，請人工核對："),
    ("promote_not_waitlisted", "以下 {n} 筆資料有轉正標記，但不是候補："),
    ("onsite_over_50", "現場號碼超過 50："),
]


def _messages(table, issues, kind):
    return [
        {"kind": kind, "text": title.format(n=len(issues[key])), "items": issues[key]}
        for key, title in table
        if issues[key]
    ]


def load(path):
    """讀取原始檔並完成狀態判斷與編號，回傳摘要（JSON 字串）。

    ok 為 False 時不准產生報表，messages 中的 error 說明原因。
    """
    _state.clear()
    try:
        header, rows, row_nos = read_source(path)
    except SourceFormatError as e:
        return json.dumps({
            "ok": False, "messages": [{"kind": "error", "text": str(e), "items": []}],
        }, ensure_ascii=False)

    result = build_students(header, rows, row_nos)
    errors = _messages(BLOCKER_MESSAGES, result["blockers"], "error")
    if not errors:
        assign_codes(result)
    students = result["students"]
    status_counts = {
        status: sum(1 for s in students if s.status == status)
        for status in (WAITLISTED, PROMOTED, CANCELLED)
    }
    if not errors:
        _state["students_all"] = {"header": header, "rows": rows, **result}
    return json.dumps({
        "ok": not errors,
        "total": len(rows),
        "included": sum(1 for s in students if s.reportable),
        "status_counts": {k: v for k, v in status_counts.items() if v},
        "messages": errors + _messages(WARNING_MESSAGES, result["issues"], "warn"),
    }, ensure_ascii=False)


def _clean_base_name(base_name, type_name, ext):
    base_name = base_name.strip()
    if base_name.lower().endswith(f".{ext}"):
        base_name = base_name[: -(len(ext) + 1)]
    base_name = INVALID_FILENAME_CHARS.sub("_", base_name).strip()
    return base_name if base_name.strip("_ ") else type_name


def generate(type_key, session_key, base_name):
    """產生指定類型的檔案，回傳檔案路徑清單與提示（JSON 字串）。

    產生多個檔案時，額外打包一個 zip 方便一次下載。
    """
    type_name, ext, module = OUTPUT_TYPES[type_key]
    sessions_wanted = SESSION_CHOICES[session_key]
    base_name = _clean_base_name(base_name, type_name, ext)

    shutil.rmtree(OUT_DIR, ignore_errors=True)
    os.makedirs(OUT_DIR)

    students_all = _state["students_all"]
    out_paths = module.build(students_all, sessions_wanted, OUT_DIR, base_name)

    notices = []
    if module is output_b:
        groups = group_students(students_all["students"], sessions_wanted)
        for session in sorted(sessions_wanted):
            for gender in ("男", "女"):
                if not groups.get((session, gender)):
                    notices.append(f"{group_label(session, gender)}沒有學員，未產生檔案。")

    zip_path = None
    if len(out_paths) > 1:
        zip_path = os.path.join(OUT_DIR, f"{base_name}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in out_paths:
                zf.write(p, os.path.basename(p))

    return json.dumps({"files": out_paths, "zip": zip_path, "notices": notices}, ensure_ascii=False)
