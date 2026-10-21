"""報名資料轉換程式 - 命令列互動介面"""
import os
import sys
import traceback

if sys.platform == "win32":
    import ctypes
    try:
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:
        pass
    for stream_name in ("stdin", "stdout", "stderr"):
        stream = getattr(sys, stream_name)
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

import output_a
import output_b
import output_c
import output_d
from processing import build_students, assign_codes
from reader import read_source

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


def base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def ask(prompt):
    return input(prompt).strip()


def main():
    folder = base_dir()
    print("========================================")
    print(" 報名資料轉換程式")
    print("========================================")
    print(f"工作資料夾：{folder}\n")

    src_name = ask("請輸入原始 Excel 檔名（不需副檔名，需與本程式放在同一資料夾）：")
    src_path = os.path.join(folder, src_name)
    if not os.path.isfile(src_path):
        for ext in (".xlsx", ".xls"):
            candidate = os.path.join(folder, src_name + ext)
            if os.path.isfile(candidate):
                src_path = candidate
                break
    if not os.path.isfile(src_path):
        print(f"找不到檔案：{src_path}")
        return

    print("\n讀取原始資料中...")
    header, rows = read_source(src_path)
    students, conflict_rows = build_students(header, rows)
    assign_codes(students)
    print(f"共讀取 {len(rows)} 筆原始資料，其中 {len(students)} 筆可辨識梯次/性別。")
    if conflict_rows:
        print(f"警告：有 {len(conflict_rows)} 筆資料的編號類別與實際性別/梯次不符，"
              f"已在「學員整理」的原檔分頁以黃色標示，未納入後續分類與報表，請人工核對。")

    students_all = {
        "header": header, "rows": rows, "students": students,
        "conflict_rows": conflict_rows,
    }

    while True:
        print("\n請選擇要製作的檔案類型：")
        for key, (name, ext, _) in OUTPUT_TYPES.items():
            print(f"  {key}. {name}（.{ext}）")
        type_key = ask("請輸入選項編號：")
        if type_key not in OUTPUT_TYPES:
            print("無效的選項。")
            return
        type_name, ext, module = OUTPUT_TYPES[type_key]

        print("\n請選擇要製作哪一梯：")
        print("  1. 第一梯")
        print("  2. 第二梯")
        print("  3. 兩梯都要")
        session_key = ask("請輸入選項編號：")
        if session_key not in SESSION_CHOICES:
            print("無效的選項。")
            return
        sessions_wanted = SESSION_CHOICES[session_key]

        out_name = ask(f"\n請輸入輸出檔名（不需副檔名，將自動加上 .{ext}）：")
        if out_name.lower().endswith(f".{ext}"):
            out_name = out_name[: -(len(ext) + 1)]

        print(f"產生「{type_name}」檔案中...")
        out_paths = module.build(students_all, sessions_wanted, folder, out_name)

        print(f"\n完成！共產生 {len(out_paths)} 個檔案：")
        for p in out_paths:
            print(f"  {p}")

        again = ask("\n是否再次生成檔案？(Y/N)：")
        if again.lower() not in ("y", "yes"):
            break


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("\n發生錯誤：")
        traceback.print_exc()
    input("\n按 Enter 鍵結束...")
