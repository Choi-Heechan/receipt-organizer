"""Scan a receipt root folder and list what Claude should read."""
import json
import os
import re
import sys
from pathlib import Path

TARGET_EXTS = {".jpg", ".jpeg", ".png", ".pdf"}
IGNORED_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}
_OUTPUT_RE = re.compile(
    r"^(receipt_.+_\d{8}(_\d+)?\.xlsx|list_receipt_\d{8}(_\d+)?\.xlsx"
    r"|receipt_\d{8}(_\d+)?\.json|dashboard_\d{8}(_\d+)?\.html)$")


def _ignored(name):
    return name.startswith((".", "~$")) or name.lower() in IGNORED_NAMES


def _is_target(name):
    return os.path.splitext(name)[1].lower() in TARGET_EXTS


def _scan_person(folder, root, skip_outputs=False):
    """Return (files relative to folder, skipped relative to root)."""
    files, skipped = [], []
    for dirpath, dirnames, filenames in os.walk(str(folder)):
        dirnames[:] = sorted(d for d in dirnames if not _ignored(d))
        for fn in filenames:
            if _ignored(fn):
                continue
            if skip_outputs and Path(dirpath) == folder and _OUTPUT_RE.match(fn):
                continue
            full = Path(dirpath) / fn
            if _is_target(fn):
                files.append(full.relative_to(folder).as_posix())
            else:
                skipped.append({"path": full.relative_to(root).as_posix(),
                                "reason": "unsupported"})
    files.sort()
    skipped.sort(key=lambda s: s["path"])
    return files, skipped


def scan(root):
    root = Path(root).resolve()
    subdirs = sorted((p for p in root.iterdir() if p.is_dir() and not _ignored(p.name)),
                     key=lambda p: p.name)
    result = {"root": str(root), "mode": "multi" if subdirs else "single",
              "folders": [], "empty_folders": [], "skipped": [], "loose_files": []}
    if subdirs:
        targets = subdirs
        result["loose_files"] = sorted(
            p.name for p in root.iterdir()
            if p.is_file() and not _ignored(p.name) and _is_target(p.name))
    else:
        targets = [root]
    for folder in targets:
        files, skipped = _scan_person(folder, root, skip_outputs=not subdirs)
        result["folders"].append({"name": folder.name, "path": str(folder),
                                  "files": files})
        result["skipped"].extend(skipped)
        if not files:
            result["empty_folders"].append(folder.name)
    return result


def main(argv):
    if len(argv) != 1:
        sys.stderr.write("使い方: scan_folder.py <フォルダ>\n")
        return 1
    root = Path(argv[0])
    if not root.is_dir():
        sys.stderr.write("フォルダが見つかりません: %s\n" % argv[0])
        return 1
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stdout.write(json.dumps(scan(root), ensure_ascii=False, indent=2) + "\n")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
