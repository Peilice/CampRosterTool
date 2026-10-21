"""將原始報名資料轉換為結構化的學員紀錄，並計算狀態、梯次分類與編號。

規則依據 docs/功能規格.md 第 3–5 節。
"""
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

EMPTY_NOTE_VALUES = {"", "無", "--"}

# 學員狀態（功能規格 §4.2）
ADMITTED = "正取"
WAITLISTED = "候補"
PROMOTED = "轉正"
CANCELLED = "取消"
REPORTABLE_STATUSES = {ADMITTED, PROMOTED}

# 管理者註記事項開頭的標記；「取消」後面可再接原本的「轉正N」以保留號碼
_MARKER_RE = re.compile(r"^(取消)?\s*(?:轉正([1-9]\d*))?")

_FULLWIDTH_SEPARATORS = str.maketrans("／－．", "/-.")


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
    return (text or "").strip() in EMPTY_NOTE_VALUES


def clean_note(text):
    """備註欄位的空白值（空字串、「無」、「--」）一律轉為空字串。"""
    return "" if is_blank_note(text) else text.strip()


def parse_marker(note):
    """回傳 (是否取消, 轉正 N 或 None)。"""
    m = _MARKER_RE.match((note or "").strip())
    cancelled = m.group(1) is not None
    promote_n = int(m.group(2)) if m.group(2) else None
    return cancelled, promote_n


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
        "raw", "row_no", "session", "gender", "source_type", "waitlisted",
        "status", "promote_n", "seq", "code",
    )

    def __init__(self, raw, row_no, session, gender, source_type):
        self.raw = raw
        self.row_no = row_no
        self.session = session
        self.gender = gender
        self.source_type = source_type
        self.waitlisted = "候補" in raw.get("參加梯次", "")
        self.status = ADMITTED
        self.promote_n = None
        self.seq = None
        self.code = None

    def get(self, key):
        return self.raw.get(key, "")

    @property
    def prefix(self):
        return SESSION_PREFIX[(self.session, self.gender)]

    @property
    def reportable(self):
        return self.status in REPORTABLE_STATUSES

    @property
    def label(self):
        """訊息中指稱這筆資料的方式，例如「第 812 列 王小明」。"""
        return f"第 {self.row_no} 列 {self.get('學員姓名').strip()}"

    @property
    def birth_ymd(self):
        y = self.get("出生年").strip()
        m = self.get("月").strip().zfill(2)
        d = self.get("日").strip().zfill(2)
        return f"{y}/{m}/{d}"


def group_label(session, gender):
    return f"第{SESSION_LABEL[session]}梯{gender}"


def _valid_ymd(y, m, d):
    return (
        y.isdigit() and 1 <= len(y) <= 3 and int(y) > 0
        and m.isdigit() and 1 <= int(m) <= 12
        and d.isdigit() and 1 <= int(d) <= 31
    )


def _normalize_birth(raw):
    """依功能規格 §3.5 解讀出生年月日，成功時改寫 raw，回傳是否可解讀。

    全部空白視為未填，不算無法解讀。
    """
    year = (raw.get("出生年") or "").strip()
    month = (raw.get("月") or "").strip()
    day = (raw.get("日") or "").strip()
    if not (year or month or day):
        return True

    y, m, d = year, month, day
    if not month and not day:
        combined = year.translate(_FULLWIDTH_SEPARATORS)
        parts = [p.strip() for p in re.split(r"[/\-.]", combined) if p.strip()]
        if len(parts) == 3:
            y, m, d = parts
        elif re.fullmatch(r"\d{7}", combined):
            y, m, d = combined[:3], combined[3:5], combined[5:]
    if re.fullmatch(r"\d{4}", y) and int(y) > 1911:
        y = str(int(y) - 1911)

    if not _valid_ymd(y, m, d):
        return False
    raw["出生年"], raw["月"], raw["日"] = y, m, d
    return True


def _onsite_code(raw):
    """現場報名的項次即為 6 碼完整編號；格式不符時回傳 None。

    試算表常把 000012 這類編號當成數字而去掉開頭的 0，不到 6 位時補回。
    """
    item = (raw.get("項次") or "").strip()
    return item.zfill(6) if re.fullmatch(r"\d{1,6}", item) else None


def _has_code_conflict(raw, session, gender, source_type):
    """判斷現場報名編號的組別碼（第 1 碼）是否與實際梯次、性別不符。"""
    code = _onsite_code(raw) if source_type == "現場" else None
    return code is not None and not code.startswith(SESSION_PREFIX[(session, gender)])


def _unrecognized_reason(raw, session, gender, source_type):
    reasons = []
    if session is None:
        value = raw.get("參加梯次", "").strip()
        reasons.append(f"參加梯次「{value}」無法辨識" if value else "參加梯次空白")
    if gender not in ("男", "女"):
        reasons.append(f"性別「{gender}」無法辨識" if gender else "性別空白")
    return "、".join(reasons)


def _assign_status(s, issues):
    cancelled, promote_n = parse_marker(s.get("管理者註記事項"))
    s.promote_n = promote_n if s.waitlisted else None
    if cancelled:
        s.status = CANCELLED
    elif s.waitlisted:
        s.status = PROMOTED if promote_n else WAITLISTED
    else:
        s.status = ADMITTED
        if promote_n:
            issues["promote_not_waitlisted"].append(f"{s.label}：有轉正標記但不是候補，視為正取")


def _item_no(s):
    try:
        return int(s.get("項次").strip())
    except (TypeError, ValueError):
        return None


def _sort_item_no(s):
    item_no = _item_no(s)
    return item_no if item_no is not None else 0


def _find_duplicates(students, key):
    """回傳 key 值相同的學員清單（每組至少兩人）。key 回傳 None 者略過。"""
    seen = {}
    for s in students:
        k = key(s)
        if k is not None:
            seen.setdefault(k, []).append(s)
    return [group for group in seen.values() if len(group) > 1]


def build_students(header, rows, row_nos):
    """依功能規格 §4 辨識學員並判斷狀態。

    回傳 dict：
      students：有效學員（含候補、取消），依原始順序排列
      unrecognized_rows / conflict_rows：被排除資料列在 rows 中的索引
      issues：各類需要人工處理的情況，每項為訊息字串清單；
              blockers 中任一項有內容時，不准產生報表（§7.3）
    """
    students = []
    unrecognized_rows = set()
    conflict_rows = set()
    issues = {
        "unrecognized": [], "conflict": [], "duplicate_id": [], "bad_birth": [],
        "promote_not_waitlisted": [], "onsite_over_50": [],
    }
    blockers = {"invalid_onsite_no": [], "duplicate_onsite_no": [], "duplicate_promote_n": []}

    for i, (row, row_no) in enumerate(zip(rows, row_nos)):
        raw = dict(zip(header, row))
        session = detect_session(raw.get("參加梯次", ""))
        gender = raw.get("性別", "").strip()
        source_type = detect_source_type(raw.get("來源", ""))
        name = raw.get("學員姓名", "").strip()
        if session is None or gender not in ("男", "女"):
            unrecognized_rows.add(i)
            reason = _unrecognized_reason(raw, session, gender, source_type)
            issues["unrecognized"].append(f"第 {row_no} 列 {name}：{reason}")
            continue
        if _has_code_conflict(raw, session, gender, source_type):
            conflict_rows.add(i)
            issues["conflict"].append(
                f"第 {row_no} 列 {name}：編號 {_onsite_code(raw)} 與{group_label(session, gender)}不符"
            )
            continue
        birth_text = "/".join(
            v for v in (raw.get(k, "").strip() for k in ("出生年", "月", "日")) if v
        )
        if not _normalize_birth(raw):
            issues["bad_birth"].append(f"第 {row_no} 列 {name}：出生年月日「{birth_text}」")
        s = Student(raw, row_no, session, gender, source_type)
        _assign_status(s, issues)
        students.append(s)

    onsite = [s for s in students if s.source_type == "現場"]
    for s in onsite:
        if _onsite_code(s.raw) is None:
            item = s.get("項次").strip()
            reason = f"項次「{item}」不是 6 碼完整編號" if item else "項次空白"
            blockers["invalid_onsite_no"].append(f"{s.label}：{reason}")
    for group in _find_duplicates(onsite, lambda s: _onsite_code(s.raw)):
        blockers["duplicate_onsite_no"].append(
            "、".join(s.label for s in group) + f"：編號都是 {_onsite_code(group[0].raw)}"
        )
    for group in _find_duplicates(students, lambda s: (s.prefix, s.promote_n) if s.promote_n else None):
        blockers["duplicate_promote_n"].append(
            "、".join(s.label for s in group) + f"：都標記為轉正{group[0].promote_n}"
        )

    def id_key(s):
        id_no = re.sub(r"\s", "", s.get("身分證字號")).upper()
        return id_no or None
    for group in _find_duplicates(students, id_key):
        issues["duplicate_id"].append("、".join(s.label for s in group) + "：身分證字號相同")

    return {
        "students": students,
        "unrecognized_rows": unrecognized_rows,
        "conflict_rows": conflict_rows,
        "issues": issues,
        "blockers": blockers,
    }


def assign_codes(result):
    """依功能規格 §5 編號並排序 result["students"]。

    須在 blockers 皆為空時呼叫（現場編號齊全且不重複、轉正 N 不重複）。
    """
    students = result["students"]
    buckets = {}
    for s in students:
        buckets.setdefault((s.prefix, s.source_type), []).append(s)

    for (prefix, source_type), group in buckets.items():
        if source_type == "現場":
            for s in group:
                if s.waitlisted and not s.promote_n:
                    continue
                s.code = _onsite_code(s.raw)
                s.seq = int(s.code[1:])
        else:
            regular = sorted((s for s in group if not s.waitlisted), key=_sort_item_no)
            promoted = sorted((s for s in group if s.promote_n), key=lambda s: s.promote_n)
            for seq, s in enumerate(regular + promoted, start=51):
                s.seq = seq
                s.code = prefix + str(seq).zfill(5)

    source_order = {"現場": 0, "網路": 1}
    students.sort(key=lambda s: (
        s.prefix, source_order[s.source_type],
        s.seq is None, s.seq if s.seq is not None else _sort_item_no(s),
    ))

    for (session, gender), prefix in SESSION_PREFIX.items():
        onsite_max = max(
            (s.seq for s in students
             if s.prefix == prefix and s.source_type == "現場" and s.seq is not None),
            default=0,
        )
        if onsite_max > 50:
            result["issues"]["onsite_over_50"].append(
                f"{group_label(session, gender)}的現場號碼最大到 {onsite_max}，超過 50，"
                f"編號會和網路報名重疊，請人工確認"
            )


def group_students(students, sessions_wanted, reportable_only=True):
    """回傳 { (session, gender): [students...] }，維持 students 的順序。

    reportable_only 為 True 時只含正取與轉正（功能規格 §4.2「其他報表」）。
    """
    groups = {}
    for s in students:
        if s.session not in sessions_wanted:
            continue
        if reportable_only and not s.reportable:
            continue
        groups.setdefault((s.session, s.gender), []).append(s)
    return groups
