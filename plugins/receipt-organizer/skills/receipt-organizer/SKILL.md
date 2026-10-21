---
name: receipt-organizer
description: |
  領収書・レシート（PDF・JPG・PNG）を個人フォルダ別に読み取り、個人別Excel・合計Excel・JSON・ダッシュボードHTMLを作成するスキル。
  上位フォルダ（その下に個人フォルダがある）のパスを指定するか、そのフォルダに移動してから実行する（상위 폴더 경로를 지정하거나 그 폴더로 이동한 뒤 실행）。
  トリガー（韓国語）：「영수증 정리」「영수증 처리」「영수증 엑셀로」
  トリガー（日本語）：「領収書処理」「領収書をまとめて」「レシートをExcelに」
  ユーザーが領収書・レシート・영수증の読み取り・整理・まとめ・Excel化を求めたら、このスキルを使うこと。
  Use when the user asks to organize receipts in a folder, e.g. 「영수증 정리해줘 C:\...\領収書_2026年10月」, or just 「영수증 처리해줘」 inside that folder.
---

# Receipt organizer

You read receipt files yourself (with the Read tool) and the bundled Python scripts turn your readings into
Excel files, a JSON file and an HTML dashboard. The plugin uses no extra server or OCR service: receipt
images are read by you (Claude) exactly as in normal Claude use, and the scripts run locally on the PC.
Follow the steps in order.

**Paths.** When this skill loads you are shown `Base directory for this skill: <path>`. That path is
`<skill dir>`; the scripts are in `<skill dir>/scripts/`. `<root>` is the folder the user named.
Always wrap every path in double quotes in commands: folder names often contain Japanese, Korean or spaces,
and on Windows they contain backslashes.

**Language.** Talk to the user in the language they used (Korean or Japanese). Fixed messages below are in
Japanese; translate them to Korean if the user writes Korean.

## Step 1 — Get the folder

- If the user gave a folder path, use it as `<root>`.
- If the user did not give a path, use the **current working directory** as `<root>` (the user may have
  moved there before asking, e.g. `cd` then 「영수증 처리해줘」). Get its absolute path with `pwd` (Git Bash /
  Mac) or `(Get-Location).Path` (PowerShell). This is tentative until confirmed in Step 3.
- Never search the disk for other receipt folders.
- Expected layout: `<root>/<person folder>/<receipt files>` (deeper subfolders count as that person's files).
  If `<root>` has no subfolders, `<root>` itself is treated as one person.

## Step 2 — Find Python 3

The scripts need Python 3 (standard library only, no pip). Find a command and call it `PY`.

- Windows (works in both Git Bash and PowerShell): run these in order and stop at the first whose output
  starts with `Python 3.`:
  1. `py -3 --version`
  2. `python --version`
  3. `python3 --version`
  A command that errors, prints nothing, opens the Microsoft Store, or prints `Python 2.` does not count.
- Mac / Linux: `python3 --version` (must start with `Python 3.`).

Use the exact working command (`py -3`, `python` or `python3`) as `PY` for every later step.

If none works, stop and tell the user (translate to Korean if needed):

> Python 3 が見つかりません。領収書の集計には Python 3（追加パッケージ不要）が必要です。

- Windows: offer to run `winget install Python.Python.3.12`. **Run it only after the user clearly says yes**
  (or let them run it themselves). If winget is not available, tell them to download Python from
  https://www.python.org/downloads/ and to tick 「Add python.exe to PATH」 on the first installer screen.
- Mac: tell them to run `xcode-select --install` (or accept the developer-tools install window macOS shows).

After installing, run the detection again. On Windows the new PATH is often not visible until a new
terminal: if detection still fails, tell the user to close and restart Claude (new session) and ask again.

## Step 3 — Scan the folder

```
PY "<skill dir>/scripts/scan_folder.py" "<root>"
```

It prints JSON:

```json
{"root": "C:\\...\\領収書_2026年10月", "mode": "multi",
 "folders": [{"name": "山田", "path": "C:\\...\\領収書_2026年10月\\山田", "files": ["20261003_taxi.jpg", "sub/a.pdf"]}],
 "empty_folders": ["佐藤"], "skipped": [{"path": "山田/x.heic", "reason": "unsupported"}], "loose_files": ["memo.pdf"]}
```

- `files` are paths relative to that folder's `path` (forward slashes).
- A folder with `files: []` also appears in `empty_folders`. **Skip folders with no files**: do not read
  them and do not put them in the readings file.
- Exit code 1 means the folder does not exist: show the message and ask for the correct path.
- If no folder has any files, tell the user there is nothing to process and stop (in current-folder mode,
  ask for the folder path instead — see below).

Tell the user briefly what you found, e.g. 「3 フォルダ・合計 24 ファイルを読み取ります。」 Keep the scan
JSON for the final report (`empty_folders`, `skipped`, `loose_files`).

**Current-folder mode (no path given in Step 1).** Before reading anything, show the absolute `<root>` and the
counts, and ask for confirmation, then wait:

> `<root>` フォルダで、個人フォルダ N 個・領収書 M 件が見つかりました。このフォルダで処理しますか？

Continue only after the user agrees. If they decline, or if no folder has any files, ask for the folder path
and restart from Step 3 with that path.

## Step 4 — Read every receipt

**Resume check (do this first).** If `<root>/.receipt_readings.json` exists, Read it. Tell the user how many
folders it already contains (e.g. 「前回の読み取り結果（2 フォルダ分）が残っています。続きから再開しますか、
最初からやり直しますか？」) and wait for the answer:
- 再開 / resume → keep its folders and skip every folder that is already complete in it (all of that
  folder's scan `files` present); read only the remaining folders.
- やり直し / start over → ignore its contents and read all folders.
Always Read the file before your first Write to it (the Write tool needs this when the file exists).

Process folders in scan order. Inside a folder, open the files **one at a time** with the Read tool
(full path = folder `path` + the relative file). Do not batch many files in one call. Assume one receipt
per file: for a multi-page PDF read only the first page(s) you need (e.g. `pages: "1"`, then `"2"` if the
total is on the next page). If a Read with `pages` fails, retry the same Read without the `pages` parameter.

Extract six fields per file:

| key | 内容 | format |
|---|---|---|
| `date` | 日付 | `YYYY/MM/DD` |
| `store` | 店名（発行者） | text as printed |
| `reg_no` | 登録番号（インボイス） | `T` + 13 digits, no spaces/hyphens |
| `amount` | 税込合計金額 | plain integer, e.g. `6050` — never `"¥6,050-"`, `"6,050円"` or `6050.0` |
| `tax_rate` | 消費税率 | `10%` / `8%` / `複数税率` |
| `payment` | 支払区分 | `現金` / `クレジット` / `PayPay` / `iD` / `QRコード` / `交通系IC` etc. |

Reading tips:
- 手書き領収証: check the stamp (印), printed parts and handwritten amount boxes carefully. The issuer is
  usually the stamp/printed name, not the 「〇〇様」 addressee.
- Receipts: information may be at the right edge, the bottom, or on a second page.
- 登録番号: look for a number starting with `T` near the store name or at the bottom, often in red or blue.
  It must be `T` + exactly 13 digits.
- 登録番号がない: if the receipt genuinely has no registration number printed (common on 手書き領収証 and
  small shops), write `reg_no` = `記載なし` — not `読み取りエラー` and not in `needs_review`.
  Use `読み取りエラー` only when a number is printed but you cannot read it.
- 元号: 令和 → 西暦 = 令和年 + 2018 (R7 = 令和7年 = 2025, R6 = 2024). 平成 + 1988.
- 金額: use the tax-included total (合計 / 領収金額 / お会計), not 小計 or お預り / お釣り.
- 税率: if both 10% and 8% items appear (e.g. 軽減税率 ※ marks), write `複数税率`.
- 支払区分: look for お預り/現金, クレジット/カード会社名, PayPay, iD, QUICPay, Suica 等. If no payment
  method is shown on the receipt (any type, handwritten or printed), write `現金`, add `payment` to
  `needs_review` and note `支払方法の記載なし`.
- 支払区分 cut off but identifiable (e.g. `三菱QR(PayPa`) → write the specific wallet (`PayPay`), add
  `payment` to `needs_review` and say so in `note`.
- 日付: if the printed date differs from the date in the file name, use the printed date, add `date` to
  `needs_review` and note it (e.g. `ファイル名の日付と異なる`).
- 店名: when the brand/shop name and the issuing company differ, write `ブランド/店名（会社名）`
  (e.g. `ローソン 新宿店（株式会社ローソン）`).
- Some receipts have rotated or sideways sections (stamps, side margins): read them too.
- Check the arithmetic: 小計 + 消費税 should equal 合計, and 税額 ≈ 税込額 × 10/110 (or 8/108).
  If ※/軽減 marks conflict with the arithmetic, the arithmetic wins: set `tax_rate` from it and add
  `tax_rate` to `needs_review` with a note.

Uncertainty rules:
- **Cannot read a field at all** → write exactly `読み取りエラー` for that field. Do not guess.
- **Read it but not sure** (blurred, overlapping, handwritten, 合計 ≠ 小計 + 税, 登録番号 digit count
  wrong, date ambiguous) → write your best value, add the field name to `needs_review`, and explain in
  `note` in short Japanese (e.g. `金額が手書きで不鮮明`, `登録番号が12桁`).
- `needs_review` is always a JSON list, using only `date`, `store`, `reg_no`, `amount`, `tax_rate`,
  `payment`; use `[]` when everything is clear. `note` is `""` when there is nothing to say.
- If the date is unreadable, still write `読み取りエラー`; the script falls back to a date at the start of
  the file name and marks it 要確認 itself.

### Readings file

After **each folder** is finished, write `<root>/.receipt_readings.json` with the Write tool, rewriting
the whole file with all folders done so far (UTF-8, no BOM, valid JSON). This keeps progress if the
session is interrupted. Format:

```json
{"folders": [
  {"folder": "山田", "receipts": [
    {"filename": "20261003_taxi.jpg", "date": "2026/10/03", "store": "東京無線タクシー",
     "reg_no": "T5130001007407", "amount": 6050, "tax_rate": "10%", "payment": "クレジット",
     "needs_review": ["amount"], "note": "金額が手書きで不鮮明"},
    {"filename": "sub/a.pdf", "date": "読み取りエラー", "store": "読み取りエラー",
     "reg_no": "読み取りエラー", "amount": "読み取りエラー", "tax_rate": "読み取りエラー",
     "payment": "読み取りエラー", "needs_review": [], "note": "画像が判読できない"}
  ]}
]}
```

- `folder` = the scan's folder `name`, exactly. `filename` = the relative path from the scan's `files`,
  exactly. One receipt entry per file, in scan order.
- Include every file of the folder, even unreadable ones (all fields `読み取りエラー`).

## Step 5 — Build the outputs

```
PY "<skill dir>/scripts/build_outputs.py" --root "<scan root>" --readings "<scan root>/.receipt_readings.json"
```

`<scan root>` is the absolute `root` value from the Step 3 scan JSON (not the text the user typed).

- Exit 0: it prints a tab-separated table (フォルダ / 件数 / 合計金額 / エラー / 要確認, last row 合計),
  then `出力:` and the full paths of the created files. The readings file is deleted automatically.
  If stderr shows `警告: 読み取り結果ファイルを削除できませんでした`, the outputs are fine; tell the user
  they may delete `.receipt_readings.json` themselves.
  Re-runs on the same day add `_2`, `_3` … instead of overwriting.
- Exit 2: a file is open (usually in Excel) and cannot be written. Show the message
  (`ファイルを閉じてから再実行してください: <path>`), ask the user to close that file, and offer to run the
  same command again. The readings file is kept, so do not read the receipts again.
- Exit 1 with `読み取り結果に含まれていないファイル:` and/or `スキャン結果にないファイル:` (lists of
  `folder/file`): nothing was written and the readings file is kept. Read each missing file (Step 4 rules),
  remove entries that are not in the scan, fix any misspelled `folder`/`filename`, Write the updated
  readings file, and run the same command again.
- Exit 1 for other reasons: show the message. If it is about the readings file (not found / JSON不正 /
  形式不正 / フォルダなし), fix `.receipt_readings.json` (valid JSON, format above) and run again.

## Step 6 — Report

Reply in the user's language with:

1. The stdout table rendered as a Markdown table, e.g.

   | フォルダ | 件数 | 合計金額 | エラー | 要確認 |
   |---|---|---|---|---|
   | 山田 | 12 | ¥54,320 | 1 | 2 |
   | 合計 | 12 | ¥54,320 | 1 | 2 |

2. The output file paths from the `出力:` lines (per-person Excel, `list_receipt_…xlsx`, `receipt_…json`,
   `dashboard_…html`).
3. A section 「処理されなかったファイル」 listing, when present: `empty_folders` (no receipt files),
   `skipped` (unsupported type such as HEIC — suggest converting to JPG/PDF), and `loose_files` (files
   placed directly in `<root>` next to person folders — suggest moving them into a person folder).
   Omit the section if all three are empty.
4. If there are エラー or 要確認 rows, say that they are shown in red / yellow in the Excel files and the
   dashboard, with the reason in the 備考 column.
5. Tell the user to open the dashboard by double-clicking `dashboard_YYYYMMDD.html` in the folder (it opens
   in the browser and works offline).
