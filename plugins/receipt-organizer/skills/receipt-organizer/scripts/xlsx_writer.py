"""Minimal stdlib-only .xlsx writer (single sheet, inline strings, named styles)."""
import re
import zipfile
from collections import namedtuple
from pathlib import Path
from xml.sax.saxutils import escape

Formula = namedtuple("Formula", "expr cached")

STYLES = ("default", "header", "normal", "normal_alt", "money", "money_alt",
          "error", "review", "review_money", "total", "total_money")

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_DOC_REL = _REL_NS + "/officeDocument"
_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

# fonts: 0 normal, 1 bold white, 2 bold red, 3 bold
_FONTS = [
    '<font><sz val="11"/><name val="Calibri"/></font>',
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',
    '<font><b/><sz val="11"/><color rgb="FFCC0000"/><name val="Calibri"/></font>',
    '<font><b/><sz val="11"/><name val="Calibri"/></font>',
]
# fills: first two are mandatory (none, gray125)
_FILL_COLORS = [None, None, "1F4E79", "D6E4F0", "FFE0E0", "FFF2A0", "FFF2CC"]
_THIN = '<{0} style="thin"><color rgb="FFBFBFBF"/></{0}>'
_BORDERS = [
    "<border><left/><right/><top/><bottom/><diagonal/></border>",
    "<border>" + "".join(_THIN.format(s) for s in ("left", "right", "top", "bottom"))
    + "<diagonal/></border>",
]
_MONEY_FMT_ID = 164

# name -> (font, fill, border, numFmt, center)
_STYLE_DEFS = {
    "default": (0, 0, 0, 0, False),
    "header": (1, 2, 1, 0, True),
    "normal": (0, 0, 1, 0, False),
    "normal_alt": (0, 3, 1, 0, False),
    "money": (0, 0, 1, _MONEY_FMT_ID, False),
    "money_alt": (0, 3, 1, _MONEY_FMT_ID, False),
    "error": (2, 4, 1, 0, False),
    "review": (0, 5, 1, 0, False),
    "review_money": (0, 5, 1, _MONEY_FMT_ID, False),
    "total": (3, 6, 1, 0, False),
    "total_money": (3, 6, 1, _MONEY_FMT_ID, False),
}

_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def _col_letter(n):
    """1 -> A, 26 -> Z, 27 -> AA."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _esc(text):
    return escape(_ILLEGAL_XML.sub("", text))


def _num(v):
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return str(int(v))
    return repr(v)


def _styles_xml():
    fills = ['<fill><patternFill patternType="none"/></fill>',
             '<fill><patternFill patternType="gray125"/></fill>']
    for rgb in _FILL_COLORS[2:]:
        fills.append('<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/>'
                     '<bgColor indexed="64"/></patternFill></fill>' % rgb)
    xfs = []
    for name in STYLES:
        font, fill, border, fmt, center = _STYLE_DEFS[name]
        attrs = 'numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"' % (
            fmt, font, fill, border)
        if fmt:
            attrs += ' applyNumberFormat="1"'
        if font:
            attrs += ' applyFont="1"'
        if fill:
            attrs += ' applyFill="1"'
        if border:
            attrs += ' applyBorder="1"'
        if center:
            xfs.append('<xf %s applyAlignment="1"><alignment horizontal="center" '
                       'vertical="center"/></xf>' % attrs)
        else:
            xfs.append("<xf %s/>" % attrs)
    return (
        _XML_DECL + '<styleSheet xmlns="%s">' % _MAIN_NS
        + '<numFmts count="1"><numFmt numFmtId="%d" formatCode="#,##0"/></numFmts>' % _MONEY_FMT_ID
        + '<fonts count="%d">%s</fonts>' % (len(_FONTS), "".join(_FONTS))
        + '<fills count="%d">%s</fills>' % (len(fills), "".join(fills))
        + '<borders count="%d">%s</borders>' % (len(_BORDERS), "".join(_BORDERS))
        + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        + '<cellXfs count="%d">%s</cellXfs>' % (len(xfs), "".join(xfs))
        + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        + "</styleSheet>"
    )


def _cell_xml(ref, value, style_idx):
    s = ' s="%d"' % style_idx
    if value is None:
        return '<c r="%s"%s/>' % (ref, s)
    if isinstance(value, Formula):
        cached = value.cached
        f = "<f>%s</f>" % _esc(value.expr)
        if cached is None:
            return '<c r="%s"%s>%s</c>' % (ref, s, f)
        if isinstance(cached, str):
            return '<c r="%s"%s t="str">%s<v>%s</v></c>' % (ref, s, f, _esc(cached))
        return '<c r="%s"%s>%s<v>%s</v></c>' % (ref, s, f, _num(cached))
    if isinstance(value, bool):
        raise ValueError("bool cell values are not supported: %s" % ref)
    if isinstance(value, (int, float)):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, _num(value))
    if isinstance(value, str):
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
            ref, s, _esc(value))
    raise ValueError("unsupported cell value type %s at %s" % (type(value).__name__, ref))


def _sheet_xml(col_widths, rows, freeze_header):
    style_index = {name: i for i, name in enumerate(STYLES)}
    parts = [_XML_DECL, '<worksheet xmlns="%s" xmlns:r="%s">' % (_MAIN_NS, _REL_NS)]
    parts.append('<sheetViews><sheetView workbookViewId="0">')
    if freeze_header:
        parts.append('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
                     '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>')
    parts.append("</sheetView></sheetViews>")
    parts.append('<sheetFormatPr defaultRowHeight="15"/>')
    if col_widths:
        parts.append("<cols>")
        for i, w in enumerate(col_widths, 1):
            parts.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (i, i, _num(w)))
        parts.append("</cols>")
    parts.append("<sheetData>")
    for r, row in enumerate(rows, 1):
        parts.append('<row r="%d">' % r)
        for c, (value, style) in enumerate(row, 1):
            if style not in style_index:
                raise ValueError("unknown style: %r" % (style,))
            parts.append(_cell_xml("%s%d" % (_col_letter(c), r), value, style_index[style]))
        parts.append("</row>")
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


def write_xlsx(path, sheet_name, col_widths, rows, freeze_header=True):
    """Write a single-sheet workbook. Each cell is (value, style_name)."""
    sheet = _sheet_xml(col_widths, rows, freeze_header)  # validates before touching disk
    content_types = (
        _XML_DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )
    root_rels = (
        _XML_DECL + '<Relationships xmlns="%s">' % _PKG_REL_NS
        + '<Relationship Id="rId1" Type="%s" Target="xl/workbook.xml"/></Relationships>' % _DOC_REL
    )
    # sheet names: max 31 chars, none of : \ / ? * [ ]
    safe_name = re.sub(r"[:\\/?*\[\]]", "_", sheet_name)[:31] or "Sheet1"
    workbook = (
        _XML_DECL + '<workbook xmlns="%s" xmlns:r="%s">' % (_MAIN_NS, _REL_NS)
        + '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>' % escape(safe_name, {'"': "&quot;"})
    )
    wb_rels = (
        _XML_DECL + '<Relationships xmlns="%s">' % _PKG_REL_NS
        + '<Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>' % _REL_NS
        + '<Relationship Id="rId2" Type="%s/styles" Target="styles.xml"/>' % _REL_NS
        + "</Relationships>"
    )
    with zipfile.ZipFile(str(Path(path)), "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", _styles_xml())
        z.writestr("xl/worksheets/sheet1.xml", sheet)
