"""Workbook health: check that formulas, drop-downs and shading rules are in place.

The expected state lives in baseline.json, made from a known-good workbook with
tools/make_baseline.py. Only formula text and drop-down/shading definitions are
stored there, never cell data.

Three kinds of check:
- Fill-down columns (Computers U-Y, Sales D/E/Q-X, Listings H): every row in the
  range holds row 2's formula with its row numbers shifted.
- One-off formulas (Summary, Selling Dashboard): each cell holds exactly the
  formula it had in the baseline.
- Drop-downs, shading rules and named ranges: each one in the baseline still
  exists with the same cells and source.
"""

from __future__ import annotations

import io
import json
import os
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

BASELINE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "baseline.json")

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
}

# A cell reference outside a string literal: optional $, column, optional $, row.
_REF = re.compile(r"(?<![A-Za-z0-9_.$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")


# --------------------------------------------------------------------------
# Formula helpers
# --------------------------------------------------------------------------

def shift_rows(formula, delta):
    """Move every relative row reference by `delta` rows (like filling down)."""
    if not delta:
        return formula
    parts = formula.split('"')
    for i in range(0, len(parts), 2):          # even parts are outside "strings"
        parts[i] = _REF.sub(
            lambda m: m.group(0) if m.group(3) else
            f"{m.group(1)}{m.group(2)}{int(m.group(4)) + delta}", parts[i])
    return '"'.join(parts)


def col_letter(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def formula_text(v):
    """openpyxl value -> formula text with '=' or None."""
    if v is None:
        return None
    text = getattr(v, "text", None)        # ArrayFormula
    if text is not None:
        v = text
    if isinstance(v, str) and v.startswith("="):
        return v
    return None


def ranges(rows):
    """[2,3,4,9] -> '2-4, 9'."""
    out, start, prev = [], None, None
    for r in sorted(rows):
        if start is None:
            start = prev = r
        elif r == prev + 1:
            prev = r
        else:
            out.append(f"{start}" if start == prev else f"{start}-{prev}")
            start = prev = r
    if start is not None:
        out.append(f"{start}" if start == prev else f"{start}-{prev}")
    return ", ".join(out)


# --------------------------------------------------------------------------
# Package (XML) reading: drop-downs, shading rules, named ranges
# --------------------------------------------------------------------------

def _sheet_parts(z):
    """{sheet name: part path} from workbook.xml and its relationships."""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels.findall("rel:Relationship", NS)}
    out = {}
    for s in wb.find("m:sheets", NS).findall("m:sheet", NS):
        t = targets.get(s.get(f"{{{NS['r']}}}id"))
        if t:
            out[s.get("name")] = t.lstrip("/") if t.startswith("/") else posixpath.normpath("xl/" + t)
    return out


def _norm_sqref(s):
    return " ".join(sorted((s or "").split()))


def read_package(data):
    """Drop-downs, shading rules and defined names straight from the XML."""
    z = zipfile.ZipFile(io.BytesIO(data))
    parts = _sheet_parts(z)
    dvs, cfs = {}, {}
    for name, part in parts.items():
        if part not in z.namelist():
            continue
        root = ET.fromstring(z.read(part))
        found = []
        for dv in root.iter(f"{{{NS['m']}}}dataValidation"):
            f1 = dv.find("m:formula1", NS)
            found.append({"type": dv.get("type"), "sqref": _norm_sqref(dv.get("sqref")),
                          "source": (f1.text or "") if f1 is not None else ""})
        for dv in root.iter(f"{{{NS['x14']}}}dataValidation"):
            f = dv.find("x14:formula1/xm:f", NS)
            sq = dv.find("xm:sqref", NS)
            found.append({"type": dv.get("type"), "sqref": _norm_sqref(sq.text if sq is not None else ""),
                          "source": (f.text or "") if f is not None else ""})
        dvs[name] = found
        rules = []
        for cf in root.findall("m:conditionalFormatting", NS):
            for rule in cf.findall("m:cfRule", NS):
                fs = [f.text or "" for f in rule.findall("m:formula", NS)]
                rules.append({"sqref": _norm_sqref(cf.get("sqref")), "type": rule.get("type"),
                              "formula": fs[0] if fs else ""})
        cfs[name] = rules
    names = {}
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    dn = wb.find("m:definedNames", NS)
    if dn is not None:
        for d in dn.findall("m:definedName", NS):
            if d.get("localSheetId") is None:
                names[d.get("name")] = d.text or ""
    return {"sheets": list(parts), "data_validations": dvs, "conditional_formats": cfs,
            "defined_names": names}


# --------------------------------------------------------------------------
# Baseline
# --------------------------------------------------------------------------

def load_baseline(path=BASELINE_PATH):
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def make_baseline(wbf, data, fill_down, fixed_sheets):
    """Record the current formulas, drop-downs, shading rules and named ranges.

    fill_down: [(sheet, header, first_row, last_row)]
    fixed_sheets: sheets whose every formula cell is recorded as-is.
    """
    pkg = read_package(data)
    base = {"version": 1, "fill_down": [], "fixed": {}, "data_validations": {},
            "conditional_formats": {}, "defined_names": pkg["defined_names"]}
    for sheet, header, first, last in fill_down:
        ws = wbf[sheet]
        hdr = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        col = hdr.index(header) + 1
        tmpl = formula_text(ws.cell(first, col).value)
        if not tmpl:
            raise ValueError(f"{sheet}!{col_letter(col)}{first} has no formula to use as the pattern")
        base["fill_down"].append({"sheet": sheet, "header": header, "column": col_letter(col),
                                  "first": first, "last": last, "formula": tmpl})
    for sheet in fixed_sheets:
        cells = {}
        for row in wbf[sheet].iter_rows():
            for c in row:
                f = formula_text(c.value)
                if f:
                    cells[c.coordinate] = f
        base["fixed"][sheet] = cells
    for sheet in pkg["data_validations"]:
        if pkg["data_validations"][sheet]:
            base["data_validations"][sheet] = pkg["data_validations"][sheet]
        if pkg["conditional_formats"][sheet]:
            base["conditional_formats"][sheet] = pkg["conditional_formats"][sheet]
    return base


# --------------------------------------------------------------------------
# Checks
# --------------------------------------------------------------------------

def check(wbf, data, baseline):
    """Compare the workbook with the baseline. Returns a JSON-ready report."""
    if not baseline:
        return {"available": False}
    sections, problems = [], []

    def add(area, checked, probs):
        sections.append({"area": area, "checked": checked, "problems": len(probs)})
        problems.extend({"area": area, **p} for p in probs)

    # ---- Fill-down formula columns ----------------------------------------
    by_sheet = {}
    for spec in baseline.get("fill_down", []):
        by_sheet.setdefault(spec["sheet"], []).append(spec)
    for sheet, specs in by_sheet.items():
        area = f"{sheet} formula columns"
        if sheet not in wbf.sheetnames:
            add(area, 0, [{"level": "serious", "where": sheet,
                           "text": f"The {sheet} tab is missing."}])
            continue
        ws = wbf[sheet]
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1), [])]
        lo = min(s["first"] for s in specs)
        hi = max(s["last"] for s in specs)
        cols = {s["column"]: s for s in specs}
        # Read only the needed block once (rows lo..hi, columns of interest).
        min_c = min(col_index(c) for c in cols)
        max_c = max(col_index(c) for c in cols)
        grid = {}
        for r_i, row in enumerate(ws.iter_rows(min_row=lo, max_row=hi, min_col=min_c, max_col=max_c,
                                               values_only=True), start=lo):
            grid[r_i] = row
        checked, probs = 0, []
        for letter, spec in cols.items():
            ci = col_index(letter)
            got_header = header[ci - 1] if ci - 1 < len(header) else None
            if got_header != spec["header"]:
                probs.append({"level": "serious", "where": f"{sheet}!{letter}1",
                              "text": f"Column {letter} header is '{got_header}', expected "
                                      f"'{spec['header']}'. Columns may have been moved or deleted."})
                continue
            missing, different = [], []
            for r in range(spec["first"], spec["last"] + 1):
                checked += 1
                row = grid.get(r, ())
                v = row[ci - min_c] if ci - min_c < len(row) else None
                got = formula_text(v)
                want = shift_rows(spec["formula"], r - spec["first"])
                if got is None:
                    missing.append(r)
                elif got != want:
                    different.append(r)
            if missing:
                probs.append({"level": "serious", "where": f"{sheet}!{letter}",
                              "text": f"{spec['header']} ({letter}): no formula in row(s) {ranges(missing)}. "
                                      f"Copy the formula down from a row that has it."})
            if different:
                probs.append({"level": "warning", "where": f"{sheet}!{letter}",
                              "text": f"{spec['header']} ({letter}): formula differs from the rest of "
                                      f"the column in row(s) {ranges(different)}. Copy the formula down "
                                      f"from a row above, or tell Claude if the change was on purpose."})
        add(area, checked, probs)

    # ---- One-off formulas (Summary, Selling Dashboard) --------------------
    for sheet, cells in baseline.get("fixed", {}).items():
        area = f"{sheet} formulas"
        if sheet not in wbf.sheetnames:
            add(area, 0, [{"level": "serious", "where": sheet, "text": f"The {sheet} tab is missing."}])
            continue
        ws = wbf[sheet]
        actual = {}
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    actual[c.coordinate] = c.value
        missing = [a for a, f in cells.items() if formula_text(actual.get(a)) is None]
        different = [a for a, f in cells.items()
                     if formula_text(actual.get(a)) is not None and formula_text(actual.get(a)) != f]
        probs = []
        if missing:
            probs.append({"level": "serious", "where": sheet,
                          "text": f"No formula in {', '.join(missing[:25])}"
                                  + (f" and {len(missing) - 25} more" if len(missing) > 25 else "")
                                  + ". These cells were formulas before."})
        if different:
            probs.append({"level": "warning", "where": sheet,
                          "text": f"Formula changed in {', '.join(different[:25])}"
                                  + (f" and {len(different) - 25} more" if len(different) > 25 else "")
                                  + ". If the change was on purpose, tell Claude to update the reference copy."})
        add(area, len(cells), probs)

    # ---- Package checks ---------------------------------------------------
    pkg = read_package(data)

    def same(a, b, keys):
        return all(a.get(k) == b.get(k) for k in keys)

    exp_dv = baseline.get("data_validations", {})
    probs, n = [], 0
    for sheet, items in exp_dv.items():
        have = pkg["data_validations"].get(sheet)
        for it in items:
            n += 1
            what = (f"{it['sqref']}: " + (f"drop-down from {it['source']}" if it["type"] == "list"
                                                  else f"{it['type']} rule"))
            if have is None:
                probs.append({"level": "serious", "where": sheet, "text": f"{what} - the tab is missing."})
            elif not any(same(it, h, ("type", "sqref", "source")) for h in have):
                alt = next((h for h in have if h["sqref"] == it["sqref"]), None)
                text = (f"{what} has changed (now {alt['type']} {alt['source']})." if alt
                        else f"{what} is missing.")
                probs.append({"level": "warning", "where": sheet, "text": text})
    add("Drop-downs and entry rules", n, probs)

    exp_cf = baseline.get("conditional_formats", {})
    probs, n = [], 0
    for sheet, items in exp_cf.items():
        have = pkg["conditional_formats"].get(sheet)
        for it in items:
            n += 1
            if have is None:
                probs.append({"level": "serious", "where": sheet,
                              "text": f"Shading rule on {it['sqref']}: the {sheet} tab is missing."})
            elif not any(same(it, h, ("sqref", "formula")) for h in have):
                alt = next((h for h in have if h["formula"] == it["formula"]), None)
                text = (f"Shading rule {it['formula']} now covers {alt['sqref']} instead of {it['sqref']}. "
                        "A paste can do this; Home > Paste > Values avoids it." if alt
                        else f"Shading rule {it['formula']} on {it['sqref']} is missing.")
                probs.append({"level": "warning", "where": sheet, "text": text})
    for name, ref in baseline.get("defined_names", {}).items():
        n += 1
        got = pkg["defined_names"].get(name)
        if got != ref:
            probs.append({"level": "warning", "where": "Name Manager",
                          "text": f"Named range {name} should be {ref}" + (f", is {got}." if got else ", is missing.")})
    add("Shading rules and named ranges", n, probs)

    return {"available": True, "sections": sections, "problems": problems,
            "checked": sum(s["checked"] for s in sections), "ok": not problems}
