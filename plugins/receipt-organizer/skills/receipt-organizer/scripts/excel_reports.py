"""Per-folder and summary Excel reports."""
from pathlib import Path

from model import ERROR
from xlsx_writer import Formula, write_xlsx

_FIELD_JA = {"date": "日付", "store": "店名", "reg_no": "登録番号",
             "amount": "金額", "tax_rate": "消費税率", "payment": "支払区分"}

_FOLDER_HEADERS = ["No.", "ファイル名", "日付", "店名", "登録番号", "金額（円）",
                   "消費税率", "支払区分", "備考"]
_FOLDER_WIDTHS = [6, 32, 14, 32, 20, 14, 14, 20, 30]
_FOLDER_FIELDS = [None, None, "date", "store", "reg_no", "amount", "tax_rate", "payment", None]

_SUMMARY_HEADERS = ["No.", "サブフォルダ名", "ファイル数", "合計金額（円）",
                    "エラー件数", "要確認件数", "出力ファイル名"]
_SUMMARY_WIDTHS = [6, 28, 12, 18, 12, 12, 40]


def _header(names):
    return [(n, "header") for n in names]


def _plain(alt, money=False):
    base = "money" if money else "normal"
    return base + "_alt" if alt else base


def remarks(receipt):
    parts = []
    review = receipt["needs_review"]
    if review:
        parts.append("要確認: " + "、".join(_FIELD_JA[f] for f in review))
    if receipt.get("note"):
        parts.append(receipt["note"])
    return "／".join(parts)


def _folder_row(receipt, alt):
    row = []
    values = [receipt["no"], receipt["filename"], receipt["date"], receipt["store"],
              receipt["reg_no"], receipt["amount"], receipt["tax_rate"],
              receipt["payment"], remarks(receipt)]
    for value, field in zip(values, _FOLDER_FIELDS):
        money = field == "amount"
        if value == ERROR:
            style = "error"
        elif field in receipt["needs_review"]:
            style = "review_money" if money else "review"
        else:
            style = _plain(alt, money)
        row.append((value, style))
    return row


def write_folder_xlsx(path, folder_summary):
    receipts = folder_summary["receipts"]
    rows = [_header(_FOLDER_HEADERS)]
    for i, r in enumerate(receipts, 1):
        rows.append(_folder_row(r, i % 2 == 0))
    total = [("合計", "total")] + [(None, "total")] * 8
    if receipts:
        total[5] = (Formula("SUM(F2:F%d)" % (len(receipts) + 1),
                            folder_summary["total_amount"]), "total_money")
    else:
        total[5] = (0, "total_money")
    rows.append(total)
    write_xlsx(Path(path), "領収書一覧", _FOLDER_WIDTHS, rows)


def write_summary_xlsx(path, report):
    folders = report["folders"]
    rows = [_header(_SUMMARY_HEADERS)]
    for i, f in enumerate(folders, 1):
        alt = i % 2 == 0
        rows.append([
            (i, _plain(alt)), (f["folder"], _plain(alt)),
            (f["file_count"], _plain(alt)), (f["total_amount"], _plain(alt, True)),
            (f["error_count"], _plain(alt)), (f["review_count"], _plain(alt)),
            (f.get("output_file", ""), _plain(alt)),
        ])
    totals = report["totals"]
    total = [("合計", "total")] + [(None, "total")] * 6
    if folders:
        last = len(folders) + 1
        for idx, col, key, style in ((2, "C", "files", "total"),
                                     (3, "D", "amount", "total_money"),
                                     (4, "E", "errors", "total"),
                                     (5, "F", "reviews", "total")):
            total[idx] = (Formula("SUM(%s2:%s%d)" % (col, col, last), totals[key]), style)
    else:
        total[2] = (0, "total")
        total[3] = (0, "total_money")
        total[4] = (0, "total")
        total[5] = (0, "total")
    rows.append(total)
    write_xlsx(Path(path), "サマリー", _SUMMARY_WIDTHS, rows)
