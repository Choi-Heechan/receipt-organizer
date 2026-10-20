import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "plugins" / "receipt-organizer" / "skills" / "receipt-organizer" / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def _sheet_root(path):
    with zipfile.ZipFile(str(path)) as z:
        return ET.fromstring(z.read("xl/worksheets/sheet1.xml"))


def read_xlsx_cells(path):
    """Return {"A1": (value_text, formula_or_None)} for every cell in sheet1."""
    cells = {}
    for c in _sheet_root(path).iter("{%s}c" % NS["m"]):
        f = c.find("m:f", NS)
        if c.get("t") == "inlineStr":
            text = "".join(t.text or "" for t in c.iter("{%s}t" % NS["m"]))
        else:
            v = c.find("m:v", NS)
            text = v.text if v is not None and v.text is not None else ""
        cells[c.get("r")] = (text, f.text if f is not None else None)
    return cells


def cell_style_index(path, ref):
    for c in _sheet_root(path).iter("{%s}c" % NS["m"]):
        if c.get("r") == ref:
            return int(c.get("s", "0"))
    raise KeyError(ref)
