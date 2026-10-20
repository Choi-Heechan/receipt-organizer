import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from helpers import SCRIPTS  # noqa: E402

from model import (  # noqa: E402
    ERROR, build_report, normalize_receipt, output_names,
    resolve_suffix, summarize_folder,
)


def rec(amount, review=()):
    return normalize_receipt(
        {"filename": "a.jpg", "date": "2025/05/01", "store": "S", "reg_no": "T1",
         "amount": amount, "tax_rate": "10%", "payment": "現金",
         "needs_review": list(review)}, 1)


class ModelTest(unittest.TestCase):
    def test_amount_variants(self):
        for v in ("¥6,050", "6,050円", 6050.0, 6050, "￥6,050 円"):
            self.assertEqual(rec(v)["amount"], 6050)
        for v in ("不明", None, "", 12.5):
            self.assertEqual(rec(v)["amount"], ERROR)

    def test_amount_trailing_marks_and_decimal_strings(self):
        for v in ("¥6,050-", "6,050円－", "6050 ※", "¥6,050- ※", "6050.0", "6,050.00", "¥6,050.00-"):
            self.assertEqual(rec(v)["amount"], 6050, v)
        for v in ("6050.5", "-", "※", "6.050.0"):
            self.assertEqual(rec(v)["amount"], ERROR, v)

    def test_needs_review_string_treated_as_list(self):
        r = normalize_receipt({"filename": "a.jpg", "amount": 100, "needs_review": "amount"}, 1)
        self.assertEqual(r["needs_review"], ["amount"])

    def test_missing_fields_become_error_and_review_excludes_errors(self):
        r = normalize_receipt({"filename": "a.jpg", "amount": 100,
                               "needs_review": ["store", "amount", "bogus", "amount"]}, 3)
        self.assertEqual(r["store"], ERROR)
        self.assertEqual(r["needs_review"], ["amount"])
        self.assertEqual(r["no"], 3)
        self.assertEqual(r["note"], "")

    def test_date_fallback_from_filename(self):
        r = normalize_receipt({"filename": "20250513_ro.jpg", "date": ""}, 1)
        self.assertEqual(r["date"], "2025/05/13")
        self.assertIn("date", r["needs_review"])
        self.assertIn("日付はファイル名から推定", r["note"])
        r = normalize_receipt({"filename": "20250513_ro.jpg", "date": None, "note": "x"}, 1)
        self.assertEqual(r["note"], "x／日付はファイル名から推定")
        r = normalize_receipt({"filename": "20251399_ro.jpg"}, 1)
        self.assertEqual(r["date"], ERROR)

    def test_date_formats(self):
        for v in ("2025-5-1", "2025.05.01", "2025/5/1"):
            r = normalize_receipt({"filename": "a.jpg", "date": v}, 1)
            self.assertEqual(r["date"], "2025/05/01")

    def test_summarize_counts(self):
        rs = [rec(6050), rec("不明"), rec(2310, [ "amount"])]
        s = summarize_folder("F", rs)
        self.assertEqual((s["total_amount"], s["error_count"], s["review_count"], s["file_count"]),
                         (8360, 1, 1, 3))
        self.assertEqual(s["folder"], "F")

    def test_all_amount_errors_total_zero(self):
        s = summarize_folder("F", [rec(None), rec("x")])
        self.assertEqual(s["total_amount"], 0)

    def test_build_report(self):
        readings = {"folders": [
            {"folder": "B", "receipts": [{"filename": "a.jpg", "amount": 100}]},
            {"folder": "A", "receipts": [{"filename": "b.jpg", "amount": 200},
                                         {"filename": "c.jpg", "amount": 300}]}]}
        rep = build_report(readings, "root", "2026-10-20 10:00")
        self.assertEqual([f["folder"] for f in rep["folders"]], ["B", "A"])
        self.assertEqual(rep["folders"][1]["receipts"][1]["no"], 2)
        self.assertEqual(rep["root_folder"], "root")
        self.assertEqual(rep["generated_at"], "2026-10-20 10:00")
        t = rep["totals"]
        self.assertEqual((t["folders"], t["files"], t["amount"]), (2, 3, 600))
        self.assertEqual(t["errors"], 3)  # missing fields on every receipt
        self.assertEqual(t["reviews"], 0)

    def test_resolve_suffix(self):
        with tempfile.TemporaryDirectory() as d:
            root = Path(d)
            self.assertEqual(resolve_suffix(root, ["株式会社 A"], "20261020"), "")
            (root / "dashboard_20261020.html").write_text("x")
            self.assertEqual(resolve_suffix(root, ["株式会社 A"], "20261020"), "_2")
            (root / "receipt_株式会社 A_20261020_2.xlsx").write_text("x")
            self.assertEqual(resolve_suffix(root, ["株式会社 A"], "20261020"), "_3")

    def test_output_names_sanitize(self):
        n = output_names(["A/B:C", "株式会社 A"], "20261020", "")
        self.assertEqual(n["folders"]["A/B:C"], "receipt_A_B_C_20261020.xlsx")
        self.assertEqual(n["folders"]["株式会社 A"], "receipt_株式会社 A_20261020.xlsx")
        self.assertEqual(n["summary"], "list_receipt_20261020.xlsx")
        self.assertEqual(n["json"], "receipt_20261020.json")
        self.assertEqual(n["dashboard"], "dashboard_20261020.html")
        n = output_names(["A"], "20261020", "_2")
        self.assertEqual(n["dashboard"], "dashboard_20261020_2.html")

    def test_names_match_scan_pattern(self):
        from scan_folder import _OUTPUT_RE
        n = output_names(["株式会社 A"], "20261020", "_2")
        for name in list(n["folders"].values()) + [n["summary"], n["json"], n["dashboard"]]:
            self.assertTrue(_OUTPUT_RE.match(name), name)


if __name__ == "__main__":
    unittest.main()
