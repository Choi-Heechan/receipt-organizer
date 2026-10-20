import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from helpers import SCRIPTS, read_xlsx_cells, cell_style_index  # noqa: E402

from xlsx_writer import Formula, STYLES, write_xlsx  # noqa: E402

ROWS = [
    [("No.", "header"), ("店名", "header")],
    [(1, "normal"), ("美濃吉 <&>", "normal")],
    [("合計", "total"), (Formula("SUM(A2:A2)", 1), "total_money")],
]


class XlsxWriterTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.path = Path(self._tmp.name) / "t.xlsx"

    def test_writes_valid_zip_with_required_parts(self):
        write_xlsx(self.path, "S", [8, 32], ROWS)
        with zipfile.ZipFile(str(self.path)) as z:
            names = z.namelist()
            for part in ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                         "xl/_rels/workbook.xml.rels", "xl/styles.xml",
                         "xl/worksheets/sheet1.xml"]:
                self.assertIn(part, names)
                ElementTree.fromstring(z.read(part))

    def test_cell_values_and_formula_cache(self):
        write_xlsx(self.path, "S", [8, 32], ROWS)
        cells = read_xlsx_cells(self.path)
        self.assertEqual(cells["B2"], ("美濃吉 <&>", None))
        self.assertEqual(cells["B3"], ("1", "SUM(A2:A2)"))
        self.assertEqual(cells["A2"], ("1", None))
        self.assertEqual(cell_style_index(self.path, "A1"), STYLES.index("header"))
        self.assertEqual(cell_style_index(self.path, "B3"), STYLES.index("total_money"))

    def test_freeze_and_widths(self):
        write_xlsx(self.path, "S", [8, 32], ROWS)
        with zipfile.ZipFile(str(self.path)) as z:
            xml = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn('ySplit="1"', xml)
        self.assertIn('state="frozen"', xml)
        self.assertIn('<col min="2" max="2" width="32"', xml)

    def test_unknown_style_raises(self):
        with self.assertRaises(ValueError):
            write_xlsx(self.path, "S", [8], [[("x", "bogus")]])

    def test_styles_xml_excel_requirements(self):
        write_xlsx(self.path, "S", [8, 32], ROWS)
        ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with zipfile.ZipFile(str(self.path)) as z:
            root = ElementTree.fromstring(z.read("xl/styles.xml"))
        fills = root.findall("m:fills/m:fill/m:patternFill", ns)
        self.assertEqual([f.get("patternType") for f in fills[:2]], ["none", "gray125"])
        self.assertEqual(len(root.findall("m:cellXfs/m:xf", ns)), len(STYLES))

    def test_column_letters_beyond_z_and_none_cells(self):
        row = [("v%d" % i, "normal") for i in range(28)] + [(None, "normal")]
        write_xlsx(self.path, "S", [10], [row], freeze_header=False)
        cells = read_xlsx_cells(self.path)
        self.assertEqual(cells["AA1"][0], "v26")
        self.assertEqual(cells["AB1"][0], "v27")
        with zipfile.ZipFile(str(self.path)) as z:
            self.assertNotIn("<pane", z.read("xl/worksheets/sheet1.xml").decode("utf-8"))


if __name__ == "__main__":
    unittest.main()
