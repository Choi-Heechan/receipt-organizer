import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from helpers import SCRIPTS  # noqa: E402

from dashboard import render_dashboard, write_dashboard  # noqa: E402
from model import build_report  # noqa: E402

TEMPLATE = (SCRIPTS / "dashboard_template.html").read_text(encoding="utf-8")


def raw(store="株式会社A", amount=1000, review=()):
    return {"filename": "a.jpg", "date": "2025/05/01", "store": store, "reg_no": "T123",
            "amount": amount, "tax_rate": "10%", "payment": "現金",
            "needs_review": list(review), "note": ""}


def report(store="株式会社A"):
    return build_report({"folders": [
        {"folder": "株式会社 A", "receipts": [raw(store), raw(store, "x", ["date"])]},
        {"folder": "空", "receipts": []},
    ]}, "root", "2026-10-20 10:00")


class DashboardTest(unittest.TestCase):
    def test_embeds_report_once_and_no_placeholder(self):
        html = render_dashboard(report())
        self.assertNotIn("__REPORT_DATA__", html)
        self.assertIn("株式会社A", html)
        self.assertEqual(TEMPLATE.count("/*__REPORT_DATA__*/null"), 1)

    def test_no_external_resources(self):
        html = render_dashboard(report())
        self.assertNotIn("http://", html)
        self.assertNotIn("https://", html)

    def test_script_breakout_escaped(self):
        html = render_dashboard(report('</script><b>"x" '))
        self.assertEqual(html.count("</script>"), TEMPLATE.count("</script>"))
        self.assertIn("<\\/script>", html)
        self.assertNotIn(" ", html)

    def test_japanese_ui_labels(self):
        html = render_dashboard(report())
        for text in ("総額", "要確認", "読み取りエラー", 'lang="ja"', "領収書なし"):
            self.assertIn(text, html)

    def test_write_dashboard_utf8(self):
        import tempfile
        with tempfile.TemporaryDirectory() as d:
            p = Path(d) / "dashboard.html"
            write_dashboard(p, report())
            self.assertIn("株式会社 A", p.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
