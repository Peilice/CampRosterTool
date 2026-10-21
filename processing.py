"""將原始報名資料轉換為結構化的學員紀錄，並計算梯次分類與編號。"""
import datetime
import re

SESSION_PREFIX = {
    (1, "男"): "0",
    (1, "女"): "1",
    (2, "男"): "2",
    (2, "女"): "3",
}

SESSION_NAME = {1: "第一梯次", 2: "第二梯次"}
SESSION_LABEL = {1: "一", 2: "二"}

EMPTY_NOTE_VALUES = {"", "無"}


def detect_session(text):
    text = text or ""
    if "第一梯次" in text:
        return 1
    if "第二梯次" in text:
        return 2
    return None


def detect_source_type(text):
    text = (text or "").strip()
    if "現場" in text:
        return "現場"
    if "網路" in text:
        return "網路"
    return None


def is_blank_note(text):
    return (text or "").strip().strip("\xa0") in EMPTY_NOTE_VALUES


def mask_name(name):
    name = (name or "").strip()
    if len(name) <= 1:
        return name
    if len(name) == 2:
        return name[0] + "○"
    return name[0] + "○" * (len(name) - 2) + name[-1]


def mask_id(id_no):
    id_no = (id_no or "").strip()
    if len(id_no) <= 6:
        return id_no
    return id_no[:3] + "****" + id_no[-3:]


def roc_year_now():
    return datetime.date.today().year - 1911


class Student:
    __slots__ = (
        "raw", "session", "gender", "source_type", "seq", "code",
    )

    def __init__(self, raw, session, gender, source_type):
        self.raw = raw
        self.session = session
        self.gender = gender
        self.source_type = source_type
        self.seq = None
        self.code = None

    def get(self, key):
        return self.raw.get(key, "")

    @property
    def birth_ymd(self):
        y = self.get("出生年").strip()
        m = self.get("月").strip().zfill(2)
        d = self.get("日").strip().zfill(2)
        return f"{y}/{m}/{d}"


def _split_combined_birth(raw):
    """現場報名常將出生年月日整串填在「出生年」欄位，需拆成年/月/日三格。"""
    year = (raw.get("出生年") or "").strip()
    month = (raw.get("月") or "").strip()
    day = (raw.get("日") or "").strip()
    if month or day or not year:
        return
    parts = [p for p in re.split(r"[/\-.]", year) if p.strip()]
    if len(parts) == 3:
        raw["出生年"], raw["月"], raw["日"] = (p.strip() for p in parts)


def _has_code_conflict(raw, session, gender):
    """判斷「現場原始編號」（已完成編號的補登資料）的類別前綴，
    是否與依實際性別/梯次欄位算出的前綴不符。"""
    preassigned = (raw.get("現場原始編號") or "").strip()
    if not preassigned:
        return False
    prefix = SESSION_PREFIX.get((session, gender))
    return prefix is not None and not preassigned.startswith(prefix)


def build_students(header, rows):
    """回傳 (students, conflict_row_indexes)。

    students：依原始順序排列的 Student 清單（僅含梯次/性別可辨識的資料列）。
    conflict_row_indexes：帶有「現場原始編號」但其類別前綴與實際性別/梯次
    不符的資料列，在 rows 中的索引（從 0 起）。這些資料列不會進入分類/編號
    流程，只在原檔分頁中標示，待人工核對。
    """
    students = []
    conflict_row_indexes = set()
    for i, row in enumerate(rows):
        raw = dict(zip(header, row))
        _split_combined_birth(raw)
        session = detect_session(raw.get("參加梯次", ""))
        gender = raw.get("性別", "").strip()
        source_type = detect_source_type(raw.get("來源", ""))
        if session is None or gender not in ("男", "女") or source_type is None:
            continue
        if _has_code_conflict(raw, session, gender):
            conflict_row_indexes.add(i)
            continue
        students.append(Student(raw, session, gender, source_type))
    return students, conflict_row_indexes


def _item_no(s):
    try:
        return int(s.get("項次"))
    except (TypeError, ValueError):
        return 0


def assign_codes(students):
    """依 (梯次,性別) 分類，各自依來源(現場/網路)分開編號，並組成 6 碼編號。

    現場報名：項次即為編號，直接沿用項次值排序、編號。
    若該筆資料帶有「現場原始編號」（例如來自另一份已經完成編號的補登
    資料，且類別前綴與實際性別/梯次相符），則直接沿用該編號作為最終編號，
    不重新計算。
    網路報名：項次僅代表報名順序，依項次排序後重新流水編號（自51起）。
    最後依 (梯次,性別) 分組，同組內先列現場報名、再列網路報名（各自依編號排序）。
    """
    buckets = {}
    for s in students:
        prefix = SESSION_PREFIX[(s.session, s.gender)]
        buckets.setdefault((prefix, s.source_type), []).append(s)

    for (prefix, source_type), group in buckets.items():
        group.sort(key=_item_no)
        if source_type == "現場":
            for s in group:
                s.seq = _item_no(s)
                preassigned = s.get("現場原始編號").strip()
                s.code = preassigned if preassigned else prefix + str(s.seq).zfill(5)
        else:
            for i, s in enumerate(group, start=51):
                s.seq = i
                s.code = prefix + str(s.seq).zfill(5)

    source_order = {"現場": 0, "網路": 1}
    students.sort(key=lambda s: (
        SESSION_PREFIX[(s.session, s.gender)], source_order[s.source_type], s.seq,
    ))

    prefixes = {SESSION_PREFIX[(s.session, s.gender)] for s in students}
    for prefix in prefixes:
        onsite_count = sum(
            1 for s in students
            if SESSION_PREFIX[(s.session, s.gender)] == prefix and s.source_type == "現場"
        )
        if onsite_count > 50:
            print(f"警告：類別前綴 {prefix} 的現場報名人數為 {onsite_count}，"
                  f"已超過 50 人，編號將與網路報名（自51起）重疊，請人工確認。")


def group_students(students, sessions_wanted):
    """回傳 { (session, gender): [students...] }，依原始順序排列。"""
    groups = {}
    for s in students:
        if s.session not in sessions_wanted:
            continue
        groups.setdefault((s.session, s.gender), []).append(s)
    return groups
