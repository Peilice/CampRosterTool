"""網頁版（Pyodide）進入點：供 index.html 呼叫，沿用命令列版的處理流程。

瀏覽器中沒有真實檔案系統，原始檔會先寫入 Pyodide 的虛擬檔案系統，
各 output 模組照常把檔案存到 OUT_DIR，再由網頁讀出並提供下載。
"""
import contextlib
import io
import json
import os
import shutil
import zipfile

from main import OUTPUT_TYPES, SESSION_CHOICES
from processing import assign_codes, build_students
from reader import read_source

OUT_DIR = "/out"

_state = {}


def load(path):
    """讀取原始檔並完成分類編號，回傳摘要（JSON 字串）。"""
    log = io.StringIO()
    with contextlib.redirect_stdout(log):
        header, rows = read_source(path)
        students, conflict_rows = build_students(header, rows)
        assign_codes(students)
    _state["students_all"] = {
        "header": header, "rows": rows, "students": students,
        "conflict_rows": conflict_rows,
    }
    warnings = [line for line in log.getvalue().splitlines() if line.strip()]
    if conflict_rows:
        warnings.insert(0, (
            f"有 {len(conflict_rows)} 筆資料的編號類別與實際性別/梯次不符，"
            f"已在「學員整理」的原檔分頁以黃色標示，未納入後續分類與報表，請人工核對。"
        ))
    return json.dumps({
        "total": len(rows),
        "recognized": len(students),
        "warnings": warnings,
    }, ensure_ascii=False)


def generate(type_key, session_key, base_name):
    """產生指定類型的檔案，回傳檔案路徑清單（JSON 字串）。

    產生多個檔案時，額外打包一個 zip 方便一次下載。
    """
    _, ext, module = OUTPUT_TYPES[type_key]
    sessions_wanted = SESSION_CHOICES[session_key]
    base_name = base_name.strip()
    if base_name.lower().endswith(f".{ext}"):
        base_name = base_name[: -(len(ext) + 1)]

    shutil.rmtree(OUT_DIR, ignore_errors=True)
    os.makedirs(OUT_DIR)

    out_paths = module.build(_state["students_all"], sessions_wanted, OUT_DIR, base_name)

    zip_path = None
    if len(out_paths) > 1:
        zip_path = os.path.join(OUT_DIR, f"{base_name}.zip")
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for p in out_paths:
                zf.write(p, os.path.basename(p))

    return json.dumps({"files": out_paths, "zip": zip_path}, ensure_ascii=False)
