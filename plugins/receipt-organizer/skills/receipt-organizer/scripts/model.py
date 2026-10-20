"""Receipt readings model: normalization, summary and output-name logic."""
import datetime
import re
from pathlib import Path

ERROR = "読み取りエラー"
REVIEW = "要確認"
FIELDS = ("date", "store", "reg_no", "amount", "tax_rate", "payment")

_DATE_RE = re.compile(r"^\s*(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})\s*$")
_STRIP_AMOUNT = re.compile(r"[¥￥,，円\s]")
_TRAILING_MARKS = re.compile(r"[\s\-－※]+$")
_BAD_NAME_CHARS = re.compile(r'[\\/:*?"<>|]')


def _valid_date(y, m, d):
    try:
        datetime.date(int(y), int(m), int(d))
    except ValueError:
        return None
    return "%04d/%02d/%02d" % (int(y), int(m), int(d))


def _norm_date(value):
    if not isinstance(value, str):
        return ERROR
    m = _DATE_RE.match(value)
    return (_valid_date(*m.groups()) if m else None) or ERROR


def _norm_amount(value):
    if isinstance(value, bool) or value is None:
        return ERROR
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value.is_integer() else ERROR
    if isinstance(value, str):
        s = _STRIP_AMOUNT.sub("", _TRAILING_MARKS.sub("", value))
        m = re.fullmatch(r"(-?\d+)(?:\.0+)?", s)
        if m:
            return int(m.group(1))
    return ERROR


def _norm_text(value):
    if value is None:
        return ERROR
    s = str(value).strip()
    return s if s else ERROR


def normalize_receipt(raw, no):
    filename = str(raw.get("filename") or "")
    out = {"no": no, "filename": filename}
    out["date"] = _norm_date(raw.get("date"))
    for f in ("store", "reg_no", "tax_rate", "payment"):
        out[f] = _norm_text(raw.get(f))
    out["amount"] = _norm_amount(raw.get("amount"))
    note = str(raw.get("note") or "")
    review = raw.get("needs_review") or []
    if isinstance(review, str):
        review = [review]
    review = [f for f in review if f in FIELDS]
    if out["date"] == ERROR:
        m = re.match(r"(\d{4})(\d{2})(\d{2})", filename)
        guessed = _valid_date(*m.groups()) if m else None
        if guessed:
            out["date"] = guessed
            review.append("date")
            note = (note + "／" if note else "") + "日付はファイル名から推定"
    seen = []
    for f in review:
        if f not in seen and out[f] != ERROR:
            seen.append(f)
    out["needs_review"] = seen
    out["note"] = note
    return out


def summarize_folder(folder, receipts):
    return {
        "folder": folder,
        "file_count": len(receipts),
        "total_amount": sum(r["amount"] for r in receipts if isinstance(r["amount"], int)),
        "error_count": sum(1 for r in receipts if any(r[f] == ERROR for f in FIELDS)),
        "review_count": sum(1 for r in receipts if r["needs_review"]),
        "receipts": receipts,
    }


def build_report(readings, root_name, generated_at):
    folders = []
    for item in readings.get("folders", []):
        receipts = [normalize_receipt(r, i) for i, r in enumerate(item.get("receipts", []), 1)]
        folders.append(summarize_folder(item["folder"], receipts))
    return {
        "generated_at": generated_at,
        "root_folder": root_name,
        "totals": {
            "folders": len(folders),
            "files": sum(f["file_count"] for f in folders),
            "amount": sum(f["total_amount"] for f in folders),
            "errors": sum(f["error_count"] for f in folders),
            "reviews": sum(f["review_count"] for f in folders),
        },
        "folders": folders,
    }


def output_names(folder_names, date, suffix):
    return {
        "folders": {
            n: "receipt_%s_%s%s.xlsx" % (_BAD_NAME_CHARS.sub("_", n), date, suffix)
            for n in folder_names
        },
        "summary": "list_receipt_%s%s.xlsx" % (date, suffix),
        "json": "receipt_%s%s.json" % (date, suffix),
        "dashboard": "dashboard_%s%s.html" % (date, suffix),
    }


def _all_names(names):
    return list(names["folders"].values()) + [names["summary"], names["json"], names["dashboard"]]


def resolve_suffix(root, folder_names, date):
    root = Path(root)
    n = 1
    while True:
        suffix = "" if n == 1 else "_%d" % n
        if not any((root / name).exists() for name in _all_names(output_names(folder_names, date, suffix))):
            return suffix
        n += 1
