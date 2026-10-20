import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from helpers import cell_style_index, read_xlsx_cells  # noqa: E402

from excel_reports import write_folder_xlsx, write_summary_xlsx  # noqa: E402
from model import ERROR, build_report, summarize_folder  # noqa: E402
from xlsx_writer import STYLES  # noqa: E402


def raw(filename, amount=1000, payment="現金", review=(), note=""):
    return {"filename": filename, "date": "2025/05/01", "store": "店", "reg_no": "T123",
            "amount": amount, "tax_rate": "10%", "payment": payment,
            "needs_review": list(review), "note": note}


def folder(name, raws):
    return build_report({"folders": [{"folder": name, "receipts": raws}]}, "root", "t")["folders"][0]


class FolderXlsxTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.path = Path(self.tmp.name) / "f.xlsx"

    def test_folder_headers_match_sample(self):
        write_folder_xlsx(self.path, folder("A", [raw("1.jpg")]))
        c = read_xlsx_cells(self.path)
        got = [c[x + "1"][0] for x in "ABCDEFGHI"]
        self.assertEqual(got, ["No.", "ファイル名", "日付", "店名", "登録番号", "金額（円）",
                               "消費税率", "支払区分", "備考"])

    def test_folder_error_and_review_cells(self):
        fs = folder("A", [raw("1.jpg", review=["amount"], note="薄い"),
                          raw("2.jpg", payment=None), raw("3.jpg")])
        write_folder_xlsx(self.path, fs)
        c = read_xlsx_cells(self.path)
        self.assertEqual(c["H3"][0], ERROR)
        self.assertEqual(cell_style_index(self.path, "H3"), STYLES.index("error"))
        self.assertEqual(cell_style_index(self.path, "F2"), STYLES.index("review_money"))
        self.assertTrue(c["I2"][0].startswith("要確認: 金額"))
        self.assertIn("薄い", c["I2"][0])
        self.assertEqual(cell_style_index(self.path, "A3"), STYLES.index("normal_alt"))
        self.assertEqual(cell_style_index(self.path, "F4"), STYLES.index("money"))

    def test_folder_total_row_formula(self):
        fs = folder("A", [raw("1.jpg"), raw("2.jpg", 500), raw("3.jpg", 250)])
        write_folder_xlsx(self.path, fs)
        c = read_xlsx_cells(self.path)
        self.assertEqual(c["A5"][0], "合計")
        self.assertEqual(c["F5"], (str(fs["total_amount"]), "SUM(F2:F4)"))
        self.assertEqual(cell_style_index(self.path, "B5"), STYLES.index("total"))

    def test_folder_all_errors_total_zero_cache(self):
        fs = folder("A", [raw("1.jpg", amount="?"), raw("2.jpg", amount=None)])
        write_folder_xlsx(self.path, fs)
        c = read_xlsx_cells(self.path)
        self.assertEqual(c["F4"], ("0", "SUM(F2:F3)"))
        self.assertEqual(cell_style_index(self.path, "F2"), STYLES.index("error"))

    def test_folder_empty(self):
        write_folder_xlsx(self.path, folder("A", []))
        c = read_xlsx_cells(self.path)
        self.assertEqual(c["F2"], ("0", None))


class SummaryXlsxTest(unittest.TestCase):
    def test_summary_rows_and_totals(self):
        def mk(name, n, total, err, rev):
            return {"folder": name, "file_count": n, "total_amount": total,
                    "error_count": err, "review_count": rev, "receipts": [],
                    "output_file": "receipt_%s_20251020.xlsx" % name}
        fa, fb = mk("株式会社A", 7, 44768, 1, 0), mk("B", 4, 72348, 1, 2)
        report = {"generated_at": "t", "root_folder": "r", "folders": [fa, fb],
                  "totals": {"folders": 2, "files": 11, "amount": 117116,
                             "errors": 2, "reviews": 2}}
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "s.xlsx"
            write_summary_xlsx(p, report)
            c = read_xlsx_cells(p)
            self.assertEqual(c["B2"][0], "株式会社A")
            self.assertEqual(c["G2"][0], "receipt_株式会社A_20251020.xlsx")
            self.assertEqual(c["D4"], ("117116", "SUM(D2:D3)"))
            self.assertEqual(c["F4"], ("2", "SUM(F2:F3)"))
            self.assertEqual(c["A4"][0], "合計")

    def test_summary_empty(self):
        report = {"folders": [], "totals": {}}
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "s.xlsx"
            write_summary_xlsx(p, report)
            self.assertEqual(read_xlsx_cells(p)["D2"], ("0", None))


if __name__ == "__main__":
    unittest.main()
