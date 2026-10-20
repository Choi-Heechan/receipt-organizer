# receipt-organizer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 상위 폴더 아래 개인 폴더의 영수증(PDF·이미지)을 Claude가 읽고, 개인별/합계 엑셀·JSON·일본어 대시보드 HTML을 만드는 Claude Code 플러그인과 그 마켓플레이스 저장소를 만든다.

**Architecture:** 판독은 Claude가 SKILL.md 지시에 따라 Read 도구로 수행하고, 결과를 `.receipt_readings.json`에 기록한다. 결정적인 작업(폴더 스캔, 정규화·집계, xlsx/json/html 생성)은 Python 표준 라이브러리 스크립트가 맡는다. 저장소 루트가 마켓플레이스이고 `plugins/receipt-organizer/`가 플러그인이다.

**Tech Stack:** Python ≥3.9 표준 라이브러리만(zipfile, xml, json, unittest), 순수 HTML/CSS/JS(외부 리소스 없음), Claude Code plugin/marketplace manifest.

**Spec:** `docs/superpowers/specs/2026-10-20-receipt-organizer-design.md`

## Global Constraints

- Python 표준 라이브러리만 사용. pip 의존성 0. Python 3.9 문법 상한(`match`, `X | Y` 타입 표기 금지).
- 모든 파일 I/O는 `encoding="utf-8"`; JSON은 `ensure_ascii=False`. 경로는 `pathlib`로 다루고 일본어·공백 포함 폴더명을 지원.
- 오류 문자열은 정확히 `読み取りエラー`, 확인 표시는 정확히 `要確認`.
- 출력 파일명(실행일 YYYYMMDD, 접미사 S = "" 또는 "_2", "_3"…): `receipt_<폴더명>_YYYYMMDD{S}.xlsx`, `list_receipt_YYYYMMDD{S}.xlsx`, `receipt_YYYYMMDD{S}.json`, `dashboard_YYYYMMDD{S}.html` — 모두 상위 폴더 바로 아래.
- 대상 확장자: `.jpg .jpeg .png .pdf`(대소문자 무시).
- 색: 머리글 1F4E79/흰 글씨, 줄무늬 D6E4F0, 오류 FFE0E0/글씨 CC0000, 要確認 FFF2A0, 합계 FFF2CC.
- 대시보드 UI 문구는 일본어. 외부 URL(`http`) 참조 금지.
- 실제 영수증·결과물은 커밋하지 않는다(테스트는 합성 데이터).
- 테스트 실행: `python3 -m unittest discover -s tests -v` (저장소 루트에서).

## Review Focus

1. 일본어·공백이 들어간 폴더명(`株式会社 A`) → 출력 파일명·JSON·HTML에 그대로, 깨짐 없이 → Task 3, Task 6 테스트.
2. Claude가 금액을 `"¥6,050"`, `"6,050円"`, `6050.0` 같은 형태로 기록 → 정수 6050으로 정규화, 숫자가 없으면 `読み取りエラー` → Task 3 테스트.
3. 개인 폴더에 `.DS_Store`, `Thumbs.db`, `~$x.xlsx`, `.~lock.*#`, 숨김 파일 → 무시하고 "건너뜀"에도 나오지 않음 → Task 2 테스트.
4. 상호명에 `</script>`·따옴표·`<` 포함 → 대시보드가 깨지거나 스크립트가 끊기지 않음 → Task 5 테스트.
5. 폴더의 모든 영수증이 금액 오류 → 합계 0, 엑셀 SUM 수식 캐시값 0, 카드 정상 표시 → Task 4 테스트.

---

## File Structure

```
.claude-plugin/marketplace.json                         마켓플레이스 정의
plugins/receipt-organizer/.claude-plugin/plugin.json     플러그인 정의
plugins/receipt-organizer/skills/receipt-organizer/
  SKILL.md                    Claude용 절차(파이썬 확인 → 스캔 → 판독 → 생성 → 보고)
  scripts/xlsx_writer.py      최소 xlsx 작성기(스타일·열너비·고정·수식+캐시값)
  scripts/scan_folder.py      폴더 스캔 → JSON(stdout)
  scripts/model.py            판독값 정규화·집계·접미사 결정
  scripts/excel_reports.py    개인별/합계 엑셀 구성
  scripts/dashboard.py        템플릿에 데이터 삽입
  scripts/dashboard_template.html  카드형 대시보드(일본어)
  scripts/build_outputs.py    CLI: readings → 4종 출력 + 보고
README.md                     설치·Python 안내(한국어·일본어)
tests/helpers.py              scripts 경로를 sys.path에 추가, xlsx 읽기 유틸
tests/test_*.py
```

---

### Task 1: xlsx 작성기

**Files:**
- Create: `plugins/receipt-organizer/skills/receipt-organizer/scripts/xlsx_writer.py`
- Create: `tests/helpers.py`, `tests/test_xlsx_writer.py`

**Interfaces:**
- Produces:
  - `Formula = namedtuple("Formula", "expr cached")` — `expr`는 `=` 없이(`"SUM(F2:F8)"`).
  - `STYLES: tuple[str, ...]` = `("default","header","normal","normal_alt","money","money_alt","error","review","review_money","total","total_money")`. `money*`는 숫자서식 `#,##0`, `error`는 FFE0E0 배경·CC0000 굵은 글씨, `review*`는 FFF2A0 배경, `total*`는 FFF2CC 굵게. `header` 1F4E79 흰 굵게 가운데. 모든 스타일(default 제외) 얇은 테두리.
  - `write_xlsx(path: Path, sheet_name: str, col_widths: list[float], rows: list[list[tuple[object, str]]], freeze_header: bool = True) -> None` — 각 셀은 `(value, style)`; value는 `str|int|float|None|Formula`. 문자열은 inline string(`t="inlineStr"`), XML 이스케이프.
  - `tests/helpers.py`: `SCRIPTS: Path`, `read_xlsx_cells(path) -> dict[str, tuple[str, str|None]]` — `{"A1": (값텍스트, 수식 or None)}`, `cell_style_index(path, ref) -> int`.

- [ ] **Step 1: 실패하는 테스트 작성** — `tests/test_xlsx_writer.py`
  - `test_writes_valid_zip_with_required_parts`: 결과 zip에 `[Content_Types].xml`, `_rels/.rels`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/styles.xml`, `xl/worksheets/sheet1.xml` 존재, 모두 `ElementTree.fromstring` 파싱 성공.
  - `test_cell_values_and_formula_cache`: rows `[[("No.","header"),("店名","header")],[(1,"normal"),("美濃吉 <&>","normal")],[("合計","total"),(Formula("SUM(A2:A2)",1),"total_money")]]` → `cells["B2"]==("美濃吉 <&>",None)`, `cells["B3"]==("1","SUM(A2:A2)")`.
  - `test_freeze_and_widths`: sheet1.xml에 `<pane ... ySplit="1" ... state="frozen"/>`, `<col min="2" max="2" width="32"`.
  - `test_unknown_style_raises`: `("x","bogus")` → `ValueError`.

- [ ] **Step 2: 실패 확인** — Run: `python3 -m unittest tests.test_xlsx_writer -v` / Expected: `ModuleNotFoundError: xlsx_writer`
- [ ] **Step 3: `write_xlsx` 구현** — styles.xml의 cellXfs 순서를 `STYLES` 순서와 일치시켜 스타일 이름 → 인덱스. 열 문자 변환은 1→A, 27→AA.
- [ ] **Step 4: 통과 확인** — 같은 명령, Expected: 4 tests OK
- [ ] **Step 5: 커밋** — `git add plugins tests && git commit -m "feat: stdlib xlsx writer"`

---

### Task 2: 폴더 스캔

**Files:**
- Create: `scripts/scan_folder.py` (이하 `scripts/` = `plugins/receipt-organizer/skills/receipt-organizer/scripts/`)
- Test: `tests/test_scan_folder.py`

**Interfaces:**
- Produces: `scan(root: Path) -> dict` 와 CLI `python scan_folder.py <root>` (결과 JSON을 stdout, UTF-8). 반환 형태:
  ```json
  {"root": "<절대경로>", "mode": "multi|single",
   "folders": [{"name": "株式会社A", "path": "<절대경로>", "files": ["sub/a.jpg", "b.pdf"]}],
   "empty_folders": ["名前"], "skipped": [{"path": "株式会社A/x.heic", "reason": "unsupported"}],
   "loose_files": ["c.jpg"]}
  ```
  - `mode="single"`: 숨김 아닌 하위 폴더가 0개 → 루트 자체가 폴더 1개(`name`=루트 폴더명), `loose_files` 없음.
  - `files`: 개인 폴더 기준 상대경로, `/` 구분, 재귀, 이름순 정렬. 폴더는 이름순.
  - 무시(어디에도 안 나옴): 이름이 `.`/`~$`/`.~lock`로 시작, `Thumbs.db`, `desktop.ini`, `.DS_Store`. 숨김 폴더는 들어가지 않음.
  - `skipped`: 개인 폴더 안 비대상 확장자 파일(루트 기준 상대경로).
  - `loose_files`: multi 모드에서 루트에 바로 놓인 **대상 확장자** 파일만(엑셀 등 출력물은 무시).
  - 루트가 없거나 폴더가 아니면 CLI는 stderr에 일본어 메시지, exit 1.

- [ ] **Step 1: 실패하는 테스트 작성** — `tempfile.TemporaryDirectory`로 합성 트리 생성.
  - `test_multi_mode_lists_people_sorted_recursive`: `B/1.JPG`, `A/sub/2.pdf`, `A/1.png` → folders 이름 `["A","B"]`, A.files `["1.png","sub/2.pdf"]`.
  - `test_ignores_system_and_lock_files`: `A/.DS_Store`, `A/Thumbs.db`, `A/~$x.xlsx`, `A/.~lock.y#`, `A/.hidden.jpg` → A.files 빈 목록, skipped 빈 목록, A는 `empty_folders`에.
  - `test_unsupported_reported_as_skipped`: `A/x.heic` → `{"path":"A/x.heic","reason":"unsupported"}`.
  - `test_loose_receipts_in_root_reported_but_outputs_ignored`: 루트 `c.jpg`, `list_receipt_20261020.xlsx` + `A/1.jpg` → `loose_files==["c.jpg"]`.
  - `test_single_mode_when_no_subfolders`: 루트에 `1.jpg`만 → mode `single`, folders[0].files `["1.jpg"]`.
  - `test_japanese_and_space_names`: `株式会社 A/領収書 1.jpg` → 이름 그대로.
- [ ] **Step 2: 실패 확인** — `python3 -m unittest tests.test_scan_folder -v` / Expected: import 실패
- [ ] **Step 3: `scan(root: Path) -> dict` 와 `main(argv: list[str]) -> int` 구현** — stdout은 `sys.stdout.reconfigure(encoding="utf-8")` 후 출력(Windows 콘솔 대비).
- [ ] **Step 4: 통과 확인** — Expected: 6 tests OK
- [ ] **Step 5: 커밋** — `git commit -m "feat: folder scanner"`

---

### Task 3: 판독값 모델(정규화·집계·접미사)

**Files:**
- Create: `scripts/model.py`
- Test: `tests/test_model.py`

**Interfaces:**
- Consumes: 없음.
- Produces:
  - 상수 `ERROR = "読み取りエラー"`, `REVIEW = "要確認"`, `FIELDS = ("date","store","reg_no","amount","tax_rate","payment")`.
  - readings 파일 형식(Claude가 작성, SKILL.md에도 동일하게 기재):
    ```json
    {"folders": [{"folder": "株式会社A", "receipts": [
      {"filename": "20250501_x.jpg", "date": "2025/05/01", "store": "...", "reg_no": "T5130001007407",
       "amount": 6050, "tax_rate": "10%", "payment": "クレジット",
       "needs_review": ["amount"], "note": "金額が不鮮明"}]}]}
    ```
  - `normalize_receipt(raw: dict, no: int) -> dict` — 반환 키: `no, filename` + FIELDS + `needs_review: list[str]`, `note: str`. 규칙:
    - 누락/빈 문자열/None 필드 → `ERROR`.
    - `amount`: int면 그대로; float이고 정수값이면 int; 문자열이면 `¥ , 円 空白`을 제거 후 정수 변환; 실패 시 `ERROR`.
    - `date`: `YYYY/MM/DD`로 맞춤(`-`·`.` 구분자 허용, 한 자리 월일 0채움). `ERROR`이고 filename이 유효한 8자리 날짜로 시작하면 그 날짜를 쓰고 `needs_review`에 `"date"` 추가, note에 `日付はファイル名から推定` 추가(기존 note와 `／`로 연결).
    - `needs_review`는 FIELDS에 있는 값만, 중복 제거, 값이 `ERROR`인 필드는 제외.
  - `summarize_folder(folder: str, receipts: list[dict]) -> dict` → `{"folder","file_count","total_amount","error_count","review_count","receipts"}`. total = int 금액 합(要確認 포함, ERROR 제외). error_count = FIELDS 중 하나라도 ERROR인 영수증 수. review_count = needs_review 비어 있지 않은 영수증 수.
  - `build_report(readings: dict, root_name: str, generated_at: str) -> dict` → 스펙 JSON 형식 `{"generated_at","root_folder","totals":{"folders","files","amount","errors","reviews"},"folders":[...]}`. 각 폴더에 이후 Task 6이 `output_file`을 채움. 폴더 순서는 readings 순서 유지.
  - `resolve_suffix(root: Path, folder_names: list[str], date: str) -> str` — 4종 파일명 중 하나라도 존재하면 `_2`, `_3`… 중 모두 비어 있는 가장 작은 것; 없으면 `""`.
  - `output_names(folder_names: list[str], date: str, suffix: str) -> dict` → `{"folders": {name: "receipt_<name>_<date><suffix>.xlsx"}, "summary": ..., "json": ..., "dashboard": ...}`. 폴더명 속 `\ / : * ? " < > |`는 `_`로 치환.

- [ ] **Step 1: 실패하는 테스트 작성**
  - `test_amount_variants`: `"¥6,050"`, `"6,050円"`, `6050.0`, `6050` → 6050; `"不明"`, `None` → ERROR.
  - `test_missing_fields_become_error_and_review_excludes_errors`: `{"filename":"a.jpg","amount":100,"needs_review":["store","amount","bogus"]}` (store 없음) → store ERROR, needs_review `["amount"]`.
  - `test_date_fallback_from_filename`: `{"filename":"20250513_ro.jpg","date":""}` → date `"2025/05/13"`, `"date" in needs_review`, note에 `日付はファイル名から推定`.
  - `test_date_formats`: `"2025-5-1"` → `"2025/05/01"`.
  - `test_summarize_counts`: 금액 [6050, ERROR, 2310(要確認)] → total 8360, error_count 1, review_count 1, file_count 3.
  - `test_all_amount_errors_total_zero`: 전부 ERROR → total 0.
  - `test_resolve_suffix`: 빈 폴더 → `""`; `dashboard_20261020.html` 존재 → `"_2"`; `_2` 파일도 존재 → `"_3"`.
  - `test_output_names_sanitize`: 폴더 `"A/B:C"` → `"receipt_A_B_C_20261020.xlsx"`; `"株式会社 A"` 그대로.
- [ ] **Step 2: 실패 확인** — `python3 -m unittest tests.test_model -v`
- [ ] **Step 3: 구현** — 위 시그니처대로 `model.py`.
- [ ] **Step 4: 통과 확인** — Expected: 8 tests OK
- [ ] **Step 5: 커밋** — `git commit -m "feat: receipt normalization and summary model"`

---

### Task 4: 엑셀 보고서

**Files:**
- Create: `scripts/excel_reports.py`
- Test: `tests/test_excel_reports.py`

**Interfaces:**
- Consumes: `write_xlsx`, `Formula` (Task 1); `summarize_folder` 결과 dict, `ERROR` (Task 3).
- Produces:
  - `write_folder_xlsx(path: Path, folder_summary: dict) -> None` — 시트명 `領収書一覧`. 열 A–I: `No. / ファイル名 / 日付 / 店名 / 登録番号 / 金額（円） / 消費税率 / 支払区分 / 備考`, 너비 `[6,32,14,32,20,14,14,20,30]`. 짝수 데이터행 `*_alt` 스타일, 금액 열 `money*`. 값이 ERROR인 칸 `error`, needs_review 필드 칸 `review`(금액이면 `review_money`). 備考 = needs_review가 있으면 `要確認: <일본어 항목명들>` + note(`／` 연결), 없으면 note. 마지막 행: A `合計`(total), F `Formula("SUM(F2:F{n})", total_amount)`(total_money), 나머지 total 빈칸.
  - `write_summary_xlsx(path: Path, report: dict) -> None` — 시트명 `サマリー`. 열 `No. / サブフォルダ名 / ファイル数 / 合計金額（円） / エラー件数 / 要確認件数 / 出力ファイル名`, 너비 `[6,28,12,18,12,12,40]`. 마지막 행 `合計` + C·D·E·F SUM(캐시값 = totals).
  - 항목명 일본어: `{"date":"日付","store":"店名","reg_no":"登録番号","amount":"金額","tax_rate":"消費税率","payment":"支払区分"}`.

- [ ] **Step 1: 실패하는 테스트 작성** (helpers.read_xlsx_cells 사용)
  - `test_folder_headers_match_sample`: A1..I1이 위 머리글과 정확히 일치.
  - `test_folder_error_and_review_cells`: 영수증2의 payment ERROR → H3 값 `読み取りエラー`이고 스타일 인덱스 = `STYLES.index("error")`; 영수증1 amount 要確認 → F2 스타일 `review_money`, I2가 `要確認: 金額`으로 시작.
  - `test_folder_total_row_formula`: 3건 → A5 `合計`, F5 수식 `SUM(F2:F4)`, 캐시값 = total_amount.
  - `test_folder_all_errors_total_zero_cache`: Review Focus 5 — F 캐시값 `"0"`.
  - `test_summary_rows_and_totals`: 폴더 2개(7건 44768 오류1 요확인0 / 4건 72348 오류1 요확인2) → B2 `株式会社A`, D4 수식 `SUM(D2:D3)` 캐시 117116, F4 캐시 2, G2 출력 파일명.
- [ ] **Step 2: 실패 확인** — `python3 -m unittest tests.test_excel_reports -v`
- [ ] **Step 3: 두 함수 구현**
- [ ] **Step 4: 통과 확인** — Expected: 5 tests OK
- [ ] **Step 5: 커밋** — `git commit -m "feat: per-folder and summary excel reports"`

---

### Task 5: 대시보드

**Files:**
- Create: `scripts/dashboard_template.html`, `scripts/dashboard.py`
- Test: `tests/test_dashboard.py`

**Interfaces:**
- Consumes: `build_report` 결과 dict (Task 3, `output_file` 포함).
- Produces: `render_dashboard(report: dict) -> str`, `write_dashboard(path: Path, report: dict) -> None`.
  - 템플릿에 정확히 한 번 나오는 `/*__REPORT_DATA__*/null`을 `json.dumps(report, ensure_ascii=False)`로 치환; 치환 문자열 안의 `</`는 `<\/`로, `\u2028`/`\u2029`는 이스케이프.
  - 화면(시안 B, 일본어): `<html lang="ja">`, `<title>領収書ダッシュボード</title>`. 헤더에 `root_folder`와 `generated_at`. 요약 줄: `総額 ¥117,116 ・ 11件 ・ 読み取りエラー 2件 ・ 要確認 2件`. 카드: 폴더명, `¥` 금액(천 단위 콤마), `n件`, 오류·要確認 건수(0이면 회색), 금액 내림차순. 카드 클릭으로 아래에 해당 폴더 표(`日付 / 店名 / 金額 / 支払区分 / 登録番号 / 消費税率 / 備考`) 펼침/접힘(한 번에 하나). ERROR 칸 빨강(FFE0E0), needs_review 칸 노랑(FFF2A0). 렌더링은 `textContent`만 사용(`innerHTML`에 데이터 삽입 금지). 라이트/다크 `prefers-color-scheme` 대응, 폭 375px에서 카드 1열.
- [ ] **Step 1: 실패하는 테스트 작성**
  - `test_embeds_report_once_and_no_placeholder`: 출력에 placeholder 없음, `"株式会社A"` 포함.
  - `test_no_external_resources`: 출력에 `http://`, `https://` 없음.
  - `test_script_breakout_escaped`: Review Focus 4 — 상호 `</script><b>"x"` → 출력에서 `</script>` 출현 횟수가 템플릿 자체의 횟수와 같고, `<\/script>` 포함.
  - `test_japanese_ui_labels`: `総額`, `要確認`, `読み取りエラー`, `lang="ja"` 포함.
- [ ] **Step 2: 실패 확인** — `python3 -m unittest tests.test_dashboard -v`
- [ ] **Step 3: 템플릿과 `render_dashboard` 구현**
- [ ] **Step 4: 통과 확인** — Expected: 4 tests OK
- [ ] **Step 5: 브라우저 확인** — 합성 report로 `/tmp` 대신 스크래치 폴더에 html 생성 후 브라우저 패널에서 열어 카드 정렬·펼침·색 표시·375px 폭을 확인.
- [ ] **Step 6: 커밋** — `git commit -m "feat: card dashboard html"`

---

### Task 6: build_outputs CLI

**Files:**
- Create: `scripts/build_outputs.py`
- Test: `tests/test_build_outputs.py`

**Interfaces:**
- Consumes: Task 3 전부, `write_folder_xlsx`/`write_summary_xlsx`(Task 4), `write_dashboard`(Task 5).
- Produces: CLI
  `python build_outputs.py --root <상위폴더> --readings <readings.json> [--date YYYYMMDD] [--keep-readings]`
  - 처리: readings 로드 → `build_report(readings, root.name, 현재시각 ISO8601 로컬 TZ)` → `resolve_suffix` → 각 폴더 `output_file` 채움 → 개인별 xlsx, 합계 xlsx, JSON(들여쓰기 2), 대시보드 작성 → 성공 시 `--keep-readings`가 없으면 readings 파일 삭제.
  - stdout(UTF-8): 사람이 읽는 표
    ```
    フォルダ	件数	合計金額	エラー	要確認
    株式会社A	7	¥44,768	1	0
    ...
    合計	11	¥117,116	2	2
    出力: <생성 파일 4종 이상 절대경로, 줄마다>
    ```
  - 종료 코드: 0 성공; 1 입력 오류(readings 없음/JSON 불량/폴더 0개 — stderr 일본어 메시지); 2 `PermissionError`(stderr에 `ファイルを閉じてから再実行してください: <경로>`).
- [ ] **Step 1: 실패하는 테스트 작성** (`subprocess.run([sys.executable, build_outputs.py, ...])`, 임시 폴더)
  - `test_end_to_end_creates_four_kinds`: 폴더 2개 readings → `receipt_株式会社 A_20261020.xlsx`(Review Focus 1: 공백 포함 이름), `receipt_B_20261020.xlsx`, `list_receipt_20261020.xlsx`, `receipt_20261020.json`, `dashboard_20261020.html` 존재; JSON의 `totals.amount` 일치, `folders[0].output_file` 일치; readings 파일 삭제됨; stdout에 `合計`.
  - `test_rerun_same_day_uses_suffix`: 같은 인자로 두 번(`--keep-readings`) → 두 번째에 `_2` 파일 4종+, 첫 번째 파일 mtime 불변.
  - `test_bad_readings_exit_1`: 존재하지 않는 경로 → returncode 1.
  - `test_json_is_utf8_not_escaped`: JSON 파일 원문에 `株式会社` 그대로(`\u` 이스케이프 아님).
- [ ] **Step 2: 실패 확인** — `python3 -m unittest tests.test_build_outputs -v`
- [ ] **Step 3: 구현**
- [ ] **Step 4: 전체 테스트** — `python3 -m unittest discover -s tests -v` / Expected: 모두 OK
- [ ] **Step 5: 커밋** — `git commit -m "feat: build_outputs CLI"`

---

### Task 7: 스킬·플러그인·마켓플레이스·README

**Files:**
- Create: `plugins/receipt-organizer/skills/receipt-organizer/SKILL.md`
- Create: `plugins/receipt-organizer/.claude-plugin/plugin.json`
- Create: `.claude-plugin/marketplace.json`
- Create: `README.md`

**Interfaces:**
- Consumes: CLI들(Task 2, 6), readings 형식(Task 3).
- 값:
  - plugin.json: `{"name":"receipt-organizer","version":"0.1.0","description":"領収書（PDF・画像）を個人フォルダ別に読み取り、Excel・JSON・ダッシュボードを作成","author":{"name":"tokyoman"}}`
  - marketplace.json: `{"name":"receipt-tools","owner":{"name":"tokyoman"},"plugins":[{"name":"receipt-organizer","source":"./plugins/receipt-organizer","description":"<plugin.json과 동일>"}]}`
- SKILL.md 내용(frontmatter `name: receipt-organizer`, `description`에 트리거: 「영수증 정리」「영수증 처리」「영수증 엑셀로」「領収書処理」「領収書をまとめて」「レシートをExcelに」 + "상위 폴더 경로를 지정" 문구). 본문 절차:
  1. **폴더 확인**: 사용자가 경로를 주지 않았으면 묻는다(자동 탐색 금지).
  2. **Python 확인**(스펙 6-2 그대로): Windows `py -3 --version` → `python --version` → `python3 --version`, 출력이 `Python 3.`로 시작하는 첫 명령을 `PY`로 사용; Mac `python3`. 없으면 일본어 안내 문구와 `winget install Python.Python.3.12`(동의 시 실행) / python.org + 「Add python.exe to PATH」 / Mac `xcode-select --install`, 설치 후 재탐지(필요 시 Claude 재시작 안내).
  3. **스캔**: `PY "<skill dir>/scripts/scan_folder.py" "<root>"` 결과를 사용자에게 폴더 수·파일 수로 짧게 알림.
  4. **판독**: 폴더별로 파일을 Read로 하나씩 열어 6항목 추출(기존 receipt-processor의 판독 요령: 손글씨·스탬프, 등록번호 T+13자리, 令和 변환 R7=2025, 세액 검산). 못 읽으면 `読み取りエラー`, 읽었지만 확신이 낮거나(흐림·겹침·손글씨, 금액≠소계+세액, 등록번호 자릿수 불일치) 하면 값 기록 + `needs_review`에 필드명 + 짧은 일본어 `note`. 폴더 하나가 끝날 때마다 `<root>/.receipt_readings.json`(Task 3 형식)을 Write로 갱신.
  5. **생성**: `PY "<skill dir>/scripts/build_outputs.py" --root "<root>" --readings "<root>/.receipt_readings.json"`. exit 2면 안내 문구 전달 후 재시도 제안.
  6. **보고**: stdout 표를 그대로 마크다운 표로 보여 주고, 생성 파일 경로, `empty_folders`·`skipped`·`loose_files`를 "処理されなかったファイル"로 나열, 대시보드를 더블클릭해 열라고 안내. 사용자의 언어(한국어/일본어)로 보고.
- README.md: 한국어·일본어 병기. 설치(`/plugin marketplace add <GitHub사용자>/<저장소>` → `/plugin install receipt-organizer@receipt-tools`), 폴더 구조 예시, 사용법(「영수증 정리해줘 C:\Users\...\領収書_2026年10月」), Python 준비, 개인정보(데이터는 PC 밖으로 나가지 않음), 업데이트 방법. 같은 PC에 기존 `receipt-processor` 스킬이 있으면 트리거가 겹치므로 비활성화 권장 문구.
- [ ] **Step 1: 4개 파일 작성**
- [ ] **Step 2: 매니페스트 검증** — Run: `claude plugin validate .` 및 `claude plugin validate plugins/receipt-organizer` / Expected: 오류 없음(경고는 검토 후 판단)
- [ ] **Step 3: 커밋** — `git commit -m "feat: skill, plugin and marketplace manifests, README"`

---

### Task 8: 실제 샘플로 종단 검증

**Files:** 없음(검증만). 결과물은 커밋하지 않는다.

- [ ] **Step 1: 샘플 복사** — `~/sample-receipts/株式会社A`, `株式会社B`만 스크래치 디렉터리의 `PDF_e2e/`로 복사(원본 폴더에 쓰지 않음).
- [ ] **Step 2: 로컬 마켓플레이스 설치** — 사용자에게 `/plugin marketplace add /path/to/receipt-organizer` → `/plugin install receipt-organizer@receipt-tools` 실행을 요청(또는 `claude --plugin-dir plugins/receipt-organizer`로 새 세션).
- [ ] **Step 3: 스킬 실행** — 「영수증 정리해줘 <PDF_e2e 경로>」. Expected: 11건 처리, 4종+2개 파일 생성.
- [ ] **Step 4: 결과 대조** — 생성된 개인별 엑셀 값을 기존 샘플 `receipt_株式会社A_20260625.xlsx`·`B`와 비교(금액·등록번호 일치 여부, 차이는 要確認 표시 여부 확인). 엑셀을 Excel/Numbers로 열어 경고 없이 열리는지, 합계 행 표시 확인.
- [ ] **Step 5: 대시보드 확인** — 브라우저 패널에서 html 열어 카드 2개·펼침·색 확인.
- [ ] **Step 6: 결과 보고** — 사용자에게 차이점과 남은 문제 보고. GitHub 공개 저장소 생성·push는 사용자가 저장소 이름과 계정을 정한 뒤 별도로 확인받고 진행.
