import contextlib
import io
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent))
from helpers import SCRIPTS  # noqa: E402

import build_outputs  # noqa: E402

CLI = SCRIPTS / "build_outputs.py"
DATE = "20261020"


def raw(amount, store="店", filename="a.jpg"):
    return {"filename": filename, "date": "2025/05/01", "store": store, "reg_no": "T123",
            "amount": amount, "tax_rate": "10%", "payment": "現金",
            "needs_review": [], "note": ""}


READINGS = {"folders": [
    {"folder": "株式会社 A", "receipts": [raw("¥6,050", "株式会社テスト"),
                                         raw(1000, filename="sub/b.pdf")]},
    {"folder": "B", "receipts": [raw("x", filename="c.png")]},
]}
FILES = ["株式会社 A/a.jpg", "株式会社 A/sub/b.pdf", "B/c.png"]


class BuildOutputsTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.root = Path(self._tmp.name)
        for rel in FILES:
            p = self.root / rel
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_bytes(b"x")
        (self.root / "empty").mkdir()
        self.readings = self.root / ".receipt_readings.json"
        self.write_readings(READINGS)

    def write_readings(self, data, bom=False):
        text = json.dumps(data, ensure_ascii=False)
        self.readings.write_text(("\ufeff" if bom else "") + text, encoding="utf-8")

    def outputs(self):
        return sorted(p.name for p in self.root.iterdir()
                      if p.is_file() and not p.name.startswith("."))

    def run_main(self, *extra):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = build_outputs.main(["--root", str(self.root), "--readings",
                                       str(self.readings), "--date", DATE] + list(extra))
        return code, out.getvalue(), err.getvalue()

    def run_cli(self, *extra, readings=None):
        return subprocess.run(
            [sys.executable, str(CLI), "--root", str(self.root),
             "--readings", str(readings or self.readings), "--date", DATE] + list(extra),
            capture_output=True, encoding="utf-8")

    def test_end_to_end_creates_four_kinds(self):
        r = self.run_cli()
        self.assertEqual(r.returncode, 0, r.stderr)
        names = ["receipt_株式会社 A_%s.xlsx" % DATE, "receipt_B_%s.xlsx" % DATE,
                 "list_receipt_%s.xlsx" % DATE, "receipt_%s.json" % DATE,
                 "dashboard_%s.html" % DATE]
        for n in names:
            self.assertTrue((self.root / n).exists(), n)
        data = json.loads((self.root / names[3]).read_text(encoding="utf-8"))
        self.assertEqual(data["totals"]["amount"], 7050)
        self.assertEqual(data["folders"][0]["output_file"], names[0])
        self.assertFalse(self.readings.exists())
        self.assertIn("合計", r.stdout)
        self.assertIn("¥7,050", r.stdout)
        self.assertIn(str(self.root / names[0]), r.stdout)

    def test_rerun_same_day_uses_suffix(self):
        self.assertEqual(self.run_cli("--keep-readings").returncode, 0)
        first = self.root / ("receipt_B_%s.xlsx" % DATE)
        m = first.stat().st_mtime_ns
        r = self.run_cli("--keep-readings")
        self.assertEqual(r.returncode, 0, r.stderr)
        for n in ["receipt_株式会社 A_%s_2.xlsx", "receipt_B_%s_2.xlsx", "list_receipt_%s_2.xlsx",
                  "receipt_%s_2.json", "dashboard_%s_2.html"]:
            self.assertTrue((self.root / (n % DATE)).exists(), n)
        self.assertEqual(first.stat().st_mtime_ns, m)
        self.assertTrue(self.readings.exists())

    def test_bad_readings_exit_1(self):
        r = self.run_cli(readings=self.root / "nope.json")
        self.assertEqual(r.returncode, 1)
        self.assertTrue(r.stderr)

    def test_invalid_json_and_empty_folders_exit_1(self):
        self.readings.write_text("{oops", encoding="utf-8")
        self.assertEqual(self.run_cli().returncode, 1)
        self.readings.write_text('{"folders": []}', encoding="utf-8")
        self.assertEqual(self.run_cli().returncode, 1)

    def test_json_is_utf8_not_escaped(self):
        self.run_cli()
        text = (self.root / ("receipt_%s.json" % DATE)).read_bytes().decode("utf-8")
        self.assertIn("株式会社", text)
        self.assertNotIn("\\u", text)

    def test_missing_file_in_readings_exit_1_no_outputs(self):
        data = json.loads(json.dumps(READINGS))
        del data["folders"][0]["receipts"][1]
        self.write_readings(data)
        r = self.run_cli()
        self.assertEqual(r.returncode, 1)
        self.assertIn("読み取り結果に含まれていないファイル:", r.stderr)
        self.assertIn("sub/b.pdf", r.stderr)
        self.assertEqual(self.outputs(), [])
        self.assertTrue(self.readings.exists())

    def test_extra_file_in_readings_exit_1(self):
        data = json.loads(json.dumps(READINGS))
        data["folders"][1]["receipts"].append(raw(1, filename="ghost.jpg"))
        data["folders"].append({"folder": "幽霊", "receipts": [raw(1)]})
        self.write_readings(data)
        r = self.run_cli()
        self.assertEqual(r.returncode, 1)
        self.assertIn("スキャン結果にないファイル:", r.stderr)
        self.assertIn("ghost.jpg", r.stderr)
        self.assertIn("幽霊", r.stderr)
        self.assertEqual(self.outputs(), [])
        self.assertTrue(self.readings.exists())

    def test_bom_readings_succeed(self):
        self.write_readings(READINGS, bom=True)
        r = self.run_cli()
        self.assertEqual(r.returncode, 0, r.stderr)

    def test_missing_root_exit_1(self):
        r = subprocess.run(
            [sys.executable, str(CLI), "--root", str(self.root / "nope"),
             "--readings", str(self.readings), "--date", DATE],
            capture_output=True, encoding="utf-8")
        self.assertEqual(r.returncode, 1)
        self.assertIn("フォルダ", r.stderr)
        self.assertNotIn("Traceback", r.stderr)

    def test_other_oserror_on_write_exit_1(self):
        with mock.patch.object(build_outputs, "write_dashboard",
                               side_effect=OSError(28, "No space left on device")):
            code, _, err = self.run_main()
        self.assertEqual(code, 1)
        self.assertIn("書き込めませんでした", err)
        self.assertTrue(self.readings.exists())

    def test_permission_error_still_exit_2(self):
        with mock.patch.object(build_outputs, "write_dashboard",
                               side_effect=PermissionError(13, "denied", "x.html")):
            code, _, err = self.run_main()
        self.assertEqual(code, 2)
        self.assertIn("ファイルを閉じてから再実行してください", err)

    def test_unlink_failure_warns_and_exit_0(self):
        with mock.patch.object(Path, "unlink", side_effect=OSError(13, "busy")):
            code, out, err = self.run_main()
        self.assertEqual(code, 0)
        self.assertIn("合計", out)
        self.assertIn("削除できませんでした", err)


if __name__ == "__main__":
    unittest.main()
