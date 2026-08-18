"""Surgical cell writes into an .xlsx, leaving the rest of the file byte-identical.

Why this exists
---------------
The Cold Call workflow writes back into the live WLCC master workbook. Saving that
workbook with openpyxl is not an option: a plain load/save round-trip of
"2026 WLCC Master Report - RE Team.xlsx" silently loses

  * all 9 dropdowns on the "WLCC Active Leads" sheet (they are stored as x14
    extension data validations, which openpyxl does not support), and
  * 15 of the 21 conditional formatting rules across the Active Leads and Cold
    Database sheets, plus the named sheet views.

None of that is anywhere near the two columns we want to touch. So instead of
re-writing the workbook, this module unzips it, edits only the target cells inside
the one sheet's XML, and re-zips every other part unchanged.

Scope of what it can write: strings, dates and numbers into existing sheets. That is
all the workflow needs. It deliberately refuses to touch a cell holding a formula.
"""

from __future__ import annotations

import io
import re
import shutil
import zipfile
from collections import Counter
from datetime import date, datetime
from pathlib import Path

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Excel's 1900 date system counts from this epoch (it is 1899-12-30, not 12-31, because
# of the deliberate leap-year bug Excel keeps for Lotus compatibility).
EPOCH_1900 = date(1899, 12, 30)
EPOCH_1904 = date(1904, 1, 1)


class PatchError(RuntimeError):
    """Raised when the workbook is not shaped the way a patch needs it to be."""


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def col_letter_to_index(letter: str) -> int:
    n = 0
    for ch in letter.strip().upper():
        if not ("A" <= ch <= "Z"):
            raise PatchError(f"Not a column letter: {letter!r}")
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col_letter(index: int) -> str:
    out = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def escape_xml_text(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        # Excel cannot store these control characters; strip rather than corrupt the file.
        .translate({c: None for c in list(range(0, 9)) + [11, 12] + list(range(14, 32))})
    )


def to_serial(value: date | datetime, date1904: bool = False) -> float:
    d = value.date() if isinstance(value, datetime) else value
    epoch = EPOCH_1904 if date1904 else EPOCH_1900
    days = (d - epoch).days
    if isinstance(value, datetime):
        secs = value.hour * 3600 + value.minute * 60 + value.second
        return days + secs / 86400
    return float(days)


# ---------------------------------------------------------------------------
# Locating the sheet part inside the package
# ---------------------------------------------------------------------------

def sheet_part_name(names: dict[str, bytes], sheet_name: str) -> str:
    """Map a visible sheet name to its 'xl/worksheets/sheetN.xml' part."""
    wb = names.get("xl/workbook.xml")
    rels = names.get("xl/_rels/workbook.xml.rels")
    if wb is None or rels is None:
        raise PatchError("This does not look like an .xlsx file (no workbook part).")

    wb_xml = wb.decode("utf-8")
    rid = None
    for tag in re.findall(r"<sheet\b[^>]*/>", wb_xml):
        name = re.search(r'\bname="([^"]*)"', tag)
        ref = re.search(r'\br:id="([^"]*)"', tag)
        if name and ref and _unescape(name.group(1)) == sheet_name:
            rid = ref.group(1)
            break
    if rid is None:
        raise PatchError(f'No sheet called "{sheet_name}" in this workbook.')

    rels_xml = rels.decode("utf-8")
    for tag in re.findall(r"<Relationship\b[^>]*/>", rels_xml):
        if re.search(r'\bId="%s"' % re.escape(rid), tag):
            target = re.search(r'\bTarget="([^"]*)"', tag)
            if not target:
                break
            t = _unescape(target.group(1)).lstrip("/")
            part = t if t.startswith("xl/") else "xl/" + t
            if part not in names:
                raise PatchError(f"Sheet part {part} is missing from the workbook.")
            return part
    raise PatchError(f'Could not resolve the part for sheet "{sheet_name}".')


def _unescape(text: str) -> str:
    return (
        text.replace("&lt;", "<").replace("&gt;", ">")
        .replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")
    )


def uses_1904(names: dict[str, bytes]) -> bool:
    wb = names.get("xl/workbook.xml", b"").decode("utf-8", "ignore")
    m = re.search(r"<workbookPr\b[^>]*>", wb)
    if not m:
        return False
    return bool(re.search(r'\bdate1904="(1|true)"', m.group(0)))


# ---------------------------------------------------------------------------
# Cell surgery
# ---------------------------------------------------------------------------

# Attributes of one tag, quote-aware and non-greedy so that a self-closing tag's trailing
# "/" is not mistaken for part of an attribute value. Getting this wrong is subtly
# destructive: a pattern that lets the attribute run swallow the "/" of "<c r="E9" s="3"/>"
# then matches on to the *next* "</c>", so writing one cell silently eats its neighbours.
_ATTRS = r'((?:"[^"]*"|[^>"])*?)(/?)>'


def _element_span(xml: str, tag: str, attr_re: str, start: int = 0) -> tuple[int, int] | None:
    """Character span of the first `<tag>` whose attributes match `attr_re`.

    Handles both `<tag .../>` and `<tag ...>...</tag>`. Safe for cells and rows because
    neither nests inside itself.
    """
    for m in re.compile(r"<%s\b%s" % (tag, _ATTRS)).finditer(xml, start):
        if not re.search(attr_re, m.group(1)):
            continue
        if m.group(2) == "/":
            return m.start(), m.end()
        close = xml.find("</%s" % tag, m.end())
        if close == -1:
            return None
        gt = xml.find(">", close)
        return m.start(), (gt + 1 if gt != -1 else len(xml))
    return None


def _iter_cells(xml: str, col: str):
    """Yield (attrs, inner_text_or_None) for every cell in one column."""
    ref_re = re.compile(r'\br="%s(\d+)"' % col)
    for m in re.compile(r"<c\b%s" % _ATTRS).finditer(xml):
        ref = ref_re.search(m.group(1))
        if not ref:
            continue
        if m.group(2) == "/":
            yield int(ref.group(1)), m.group(1), None
            continue
        close = xml.find("</c", m.end())
        yield int(ref.group(1)), m.group(1), (xml[m.end():close] if close != -1 else "")


# Built-in number format ids that Excel renders as a date and/or a time.
BUILTIN_DATE_FMTS = frozenset(
    list(range(14, 23)) + list(range(27, 37)) + [45, 46, 47] + list(range(50, 59))
)


def _is_date_format_code(code: str) -> bool:
    """True if a custom format code renders as a date or time.

    Literal text inside quotes and the escaped/bracketed sections are stripped first, so a
    currency format like _([$$-409]* #,##0_) is not mistaken for a date because of its 'd'.
    """
    stripped = re.sub(r'"[^"]*"', "", code)
    stripped = re.sub(r"\[[^\]]*\]", "", stripped)
    stripped = re.sub(r"\\.", "", stripped)
    return bool(re.search(r"[ymdhs]", stripped, re.I))


def date_style_ids(styles_xml: str) -> set[str]:
    """Indices into cellXfs whose number format renders as a date or time."""
    custom_dates = {
        int(m.group(1))
        for m in re.finditer(r'<numFmt\b[^>]*\bnumFmtId="(\d+)"[^>]*/>', styles_xml)
        if (fc := re.search(r'\bformatCode="([^"]*)"', m.group(0)))
        and _is_date_format_code(_unescape(fc.group(1)))
    }
    date_fmts = BUILTIN_DATE_FMTS | custom_dates

    block = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs\s*>", styles_xml, re.S)
    if not block:
        return set()
    out = set()
    xfs = re.findall(r"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf\s*>", block.group(1), re.S)
    for idx, xf in enumerate(xfs):
        m = re.search(r'\bnumFmtId="(\d+)"', xf)
        if m and int(m.group(1)) in date_fmts:
            out.add(str(idx))
    return out


def _column_style(sheet_xml: str, col: str, skip_row: int = 1,
                  restrict_to: set[str] | None = None) -> str | None:
    """The style id used most often in this column, so a newly created cell looks like
    its neighbours.

    Only cells that actually hold a value are counted. That distinction matters: a date
    column is typically mostly empty placeholder cells carrying the General format, and
    copying one of those would make a written date show up as a bare serial number like
    46252 instead of a date. The populated cells carry the real date format.

    `restrict_to` narrows the choice to a set of acceptable style ids - used when writing a
    date, so the style picked is one that genuinely has a date number format. That matters
    here: column D's commonest style is General, because the column also holds text dates,
    and only the less common style actually formats as a date.
    """
    populated: Counter[str] = Counter()
    empty: Counter[str] = Counter()
    for row, attrs, inner in _iter_cells(sheet_xml, col):
        if row <= skip_row:
            continue
        s = re.search(r'\bs="(\d+)"', attrs)
        if not s or (restrict_to is not None and s.group(1) not in restrict_to):
            continue
        (populated if (inner or "").strip() else empty)[s.group(1)] += 1

    for counts in (populated, empty):
        if counts:
            return counts.most_common(1)[0][0]
    return None


def _build_cell(ref: str, value, style: str | None, date1904: bool) -> str:
    attrs = f' r="{ref}"'
    if style:
        attrs += f' s="{style}"'

    if value is None or value == "":
        return f"<c{attrs}/>"

    if isinstance(value, bool):
        return f'<c{attrs} t="b"><v>{int(value)}</v></c>'

    if isinstance(value, (date, datetime)):
        serial = to_serial(value, date1904)
        text = str(int(serial)) if float(serial).is_integer() else repr(serial)
        return f"<c{attrs}><v>{text}</v></c>"

    if isinstance(value, (int, float)):
        return f"<c{attrs}><v>{value}</v></c>"

    text = str(value)
    space = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c{attrs} t="inlineStr"><is><t{space}>{escape_xml_text(text)}</t></is></c>'


def _set_cell_in_row(row_xml: str, row: int, col: str, value, style_hint: str | None,
                     date1904: bool, date_styles: set[str] | None = None) -> str:
    """Return `row_xml` with cell `col` set to `value`, inserting it in column order."""
    ref = f"{col}{row}"
    span = _element_span(row_xml, "c", r'\br="%s"' % re.escape(ref))

    if span:
        old = row_xml[span[0]:span[1]]
        if "<f" in old:
            raise PatchError(f"Cell {ref} contains a formula. Refusing to overwrite it.")
        keep = re.search(r'\bs="(\d+)"', old)
        style = keep.group(1) if keep else style_hint
        # An empty cell in a date column often carries the General format. Keeping it would
        # show the date as a serial number, so a date overrides an incompatible style.
        if (
            isinstance(value, (date, datetime))
            and date_styles
            and style not in date_styles
            and style_hint in date_styles
        ):
            style = style_hint
        return row_xml[: span[0]] + _build_cell(ref, value, style, date1904) + row_xml[span[1]:]

    new_cell = _build_cell(ref, value, style_hint, date1904)
    target = col_letter_to_index(col)

    # A self-closing <row .../> has no children yet; give it a body.
    if row_xml.rstrip().endswith("/>") and "</row" not in row_xml:
        head = row_xml.rstrip()[:-2].rstrip()
        return f"{head}>{new_cell}</row>"

    # Insert before the first cell whose column sorts after ours, so cells stay in order.
    for m in re.finditer(r'<c\b[^>]*\br="([A-Z]+)\d+"', row_xml):
        if col_letter_to_index(m.group(1)) > target:
            return row_xml[: m.start()] + new_cell + row_xml[m.start():]

    close = row_xml.rfind("</row")
    if close == -1:
        raise PatchError(f"Malformed <row> for row {row}.")
    return row_xml[:close] + new_cell + row_xml[close:]


def _widen_spans(row_xml: str, cols: list[str]) -> str:
    """Keep the row's `spans` hint consistent with the cells it now holds."""
    m = re.search(r'(<row\b[^>]*?)\bspans="(\d+):(\d+)"', row_xml)
    if not m:
        return row_xml
    lo, hi = int(m.group(2)), int(m.group(3))
    idx = [col_letter_to_index(c) for c in cols]
    new_lo, new_hi = min([lo] + idx), max([hi] + idx)
    if (new_lo, new_hi) == (lo, hi):
        return row_xml
    return re.sub(r'\bspans="\d+:\d+"', f'spans="{new_lo}:{new_hi}"', row_xml, count=1)


def patch_sheet_xml(sheet_xml: str, edits: dict[int, dict[str, object]],
                    date1904: bool = False, header_row: int = 1,
                    date_styles: set[str] | None = None) -> str:
    """Apply `{row: {column_letter: value}}` to one sheet's XML string.

    `date_styles` is the set of style ids that carry a date number format, from
    `date_style_ids(styles.xml)`. It is used to pick the style for a newly created date
    cell, so the date does not land in the sheet looking like a five-digit serial number.
    """
    if not edits:
        return sheet_xml

    # A column written with a date needs a date-formatted style; anything else can take the
    # column's usual style.
    date_cols = {
        col
        for cols in edits.values()
        for col, value in cols.items()
        if isinstance(value, (date, datetime))
    }
    styles = {}
    for col in {c for cols in edits.values() for c in cols}:
        restrict = date_styles if (col in date_cols and date_styles) else None
        styles[col] = (
            _column_style(sheet_xml, col, header_row, restrict)
            # Fall back to the column's ordinary style if it has no date-formatted cell yet.
            or (_column_style(sheet_xml, col, header_row) if restrict else None)
        )

    max_written = 0
    for row in sorted(edits):
        span = _element_span(sheet_xml, "row", r'\br="%d"' % row)
        if not span:
            raise PatchError(
                f"Row {row} does not exist in the sheet, so there is nothing to update."
            )
        row_xml = sheet_xml[span[0]:span[1]]
        for col, value in edits[row].items():
            row_xml = _set_cell_in_row(
                row_xml, row, col, value, styles.get(col), date1904, date_styles
            )
        row_xml = _widen_spans(row_xml, list(edits[row]))
        sheet_xml = sheet_xml[: span[0]] + row_xml + sheet_xml[span[1]:]
        max_written = max(max_written, row)

    return _widen_dimension(sheet_xml, edits, max_written)


def _widen_dimension(sheet_xml: str, edits: dict[int, dict[str, object]], max_row: int) -> str:
    m = re.search(r'<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"\s*/>', sheet_xml)
    if not m:
        return sheet_xml
    lo_c, lo_r, hi_c, hi_r = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
    cols = [col_letter_to_index(c) for cols_ in edits.values() for c in cols_]
    new_lo = index_to_col_letter(min([col_letter_to_index(lo_c)] + cols))
    new_hi = index_to_col_letter(max([col_letter_to_index(hi_c)] + cols))
    new_ref = f"{new_lo}{min(lo_r, min(edits))}:{new_hi}{max(hi_r, max_row)}"
    if new_ref == f"{lo_c}{lo_r}:{hi_c}{hi_r}":
        return sheet_xml
    return sheet_xml[: m.start()] + f'<dimension ref="{new_ref}"/>' + sheet_xml[m.end():]


# ---------------------------------------------------------------------------
# Package level
# ---------------------------------------------------------------------------

def patch_workbook_bytes(source: bytes, sheet_name: str,
                         edits: dict[int, dict[str, object]],
                         header_row: int = 1) -> bytes:
    """Return a copy of `source` with `edits` applied to one sheet.

    `edits` is {worksheet_row: {column_letter: value}}. Every other part of the
    package - dropdowns, conditional formatting, charts, other sheets - is copied
    through untouched.
    """
    with zipfile.ZipFile(io.BytesIO(source)) as zin:
        infos = zin.infolist()
        parts = {i.filename: zin.read(i.filename) for i in infos}

    part = sheet_part_name(parts, sheet_name)
    styles = parts.get("xl/styles.xml", b"").decode("utf-8", "ignore")
    patched = patch_sheet_xml(
        parts[part].decode("utf-8"),
        edits,
        uses_1904(parts),
        header_row,
        date_style_ids(styles) if styles else None,
    )
    parts[part] = patched.encode("utf-8")

    # calcChain is a cache of formula evaluation order. Our edits are literal values, but
    # dropping it is the safe move if anything downstream recalculates: Excel rebuilds it.
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, parts[info.filename])
    return out.getvalue()


def patch_workbook_file(path: Path, sheet_name: str,
                        edits: dict[int, dict[str, object]],
                        header_row: int = 1, backup: bool = True) -> bytes:
    """Patch a workbook on disk in place, keeping a .bak alongside it by default."""
    path = Path(path)
    data = path.read_bytes()
    result = patch_workbook_bytes(data, sheet_name, edits, header_row)
    if backup:
        shutil.copy2(path, path.with_suffix(path.suffix + ".bak"))
    path.write_bytes(result)
    return result
