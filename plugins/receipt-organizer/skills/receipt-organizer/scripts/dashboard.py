"""Card-style HTML dashboard (single offline file)."""
import copy
import json
from pathlib import Path

from excel_reports import remarks

PLACEHOLDER = "/*__REPORT_DATA__*/null"


def _json_for_script(data):
    text = json.dumps(data, ensure_ascii=False)
    return text.replace("<", "\\u003c").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")


def render_dashboard(report):
    template = Path(__file__).with_name("dashboard_template.html").read_text(encoding="utf-8")
    data = copy.deepcopy(report)
    for folder in data["folders"]:
        for receipt in folder["receipts"]:
            receipt["remarks"] = remarks(receipt)
    return template.replace(PLACEHOLDER, _json_for_script(data))


def write_dashboard(path, report):
    Path(path).write_text(render_dashboard(report), encoding="utf-8")
