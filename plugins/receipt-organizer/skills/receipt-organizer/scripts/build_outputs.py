"""CLI: readings JSON -> per-folder xlsx, summary xlsx, JSON, dashboard, result table."""
import argparse
import datetime
import json
import sys
from pathlib import Path

from dashboard import write_dashboard
from excel_reports import write_folder_xlsx, write_summary_xlsx
from model import build_report, output_names, resolve_suffix


def _fail(msg, code=1):
    print(msg, file=sys.stderr)
    return code


def _table(report, paths):
    lines = ["フォルダ\t件数\t合計金額\tエラー\t要確認"]
    for f in report["folders"]:
        lines.append("%s\t%d\t¥%s\t%d\t%d" % (f["folder"], f["file_count"],
                                              format(f["total_amount"], ","),
                                              f["error_count"], f["review_count"]))
    t = report["totals"]
    lines.append("合計\t%d\t¥%s\t%d\t%d" % (t["files"], format(t["amount"], ","),
                                          t["errors"], t["reviews"]))
    lines.append("出力:")
    lines.extend(str(p) for p in paths)
    return "\n".join(lines)


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--readings", required=True)
    ap.add_argument("--date", default=datetime.date.today().strftime("%Y%m%d"))
    ap.add_argument("--keep-readings", action="store_true")
    args = ap.parse_args(argv)

    root = Path(args.root)
    readings_path = Path(args.readings)
    if not readings_path.is_file():
        return _fail("読み取り結果ファイルが見つかりません: %s" % readings_path)
    try:
        readings = json.loads(readings_path.read_text(encoding="utf-8"))
    except (ValueError, UnicodeDecodeError):
        return _fail("読み取り結果ファイルのJSONが不正です: %s" % readings_path)
    if not isinstance(readings, dict) or not readings.get("folders"):
        return _fail("読み取り結果にフォルダがありません: %s" % readings_path)

    generated_at = datetime.datetime.now().astimezone().isoformat(timespec="seconds")
    try:
        report = build_report(readings, root.name, generated_at)
    except (KeyError, TypeError, AttributeError):
        return _fail("読み取り結果の形式が不正です: %s" % readings_path)

    folder_names = [f["folder"] for f in report["folders"]]
    suffix = resolve_suffix(root, folder_names, args.date)
    names = output_names(folder_names, args.date, suffix)
    for f in report["folders"]:
        f["output_file"] = names["folders"][f["folder"]]

    paths = []
    try:
        for f in report["folders"]:
            p = root / f["output_file"]
            write_folder_xlsx(p, f)
            paths.append(p)
        p = root / names["summary"]
        write_summary_xlsx(p, report)
        paths.append(p)
        p = root / names["json"]
        p.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        paths.append(p)
        p = root / names["dashboard"]
        write_dashboard(p, report)
        paths.append(p)
    except PermissionError as e:
        return _fail("ファイルを閉じてから再実行してください: %s" % (e.filename or p), 2)

    if not args.keep_readings:
        readings_path.unlink()
    print(_table(report, [x.resolve() for x in paths]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
