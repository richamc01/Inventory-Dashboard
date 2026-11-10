"""Read Hardware Inventory.xlsx and turn it into one JSON-ready snapshot.

The workbook is read-only here. Formula columns are taken from the values
Excel saved with the file (the same numbers Excel shows). When a saved value is
missing, the same rule is recomputed in Python and a warning is added.
"""

from __future__ import annotations

import datetime as dt
import io
import os
import re
import time
import warnings
import zipfile
from collections import Counter, OrderedDict, defaultdict

import openpyxl

from . import health as health_mod

# openpyxl warns about Excel's Data Validation extension on every load.
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

READ_RETRIES = 3
RETRY_DELAY_S = 1.0


# --------------------------------------------------------------------------
# Low-level helpers
# --------------------------------------------------------------------------

def clean(v):
    """Normalise one cell value: trim text, blank -> None, dates -> ISO text."""
    if v is None:
        return None
    if isinstance(v, str):
        v = v.strip()
        return v or None
    if isinstance(v, dt.datetime):
        if v.hour == 0 and v.minute == 0 and v.second == 0:
            return v.date().isoformat()
        return v.isoformat(timespec="minutes")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def num(v):
    """Return a number or None (never raises)."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        try:
            return float(v.replace("$", "").replace(",", ""))
        except ValueError:
            return None
    return None


def natural_key(asset_id):
    """Sort AIO-2 before AIO-10."""
    s = asset_id or ""
    return [int(p) if p.isdigit() else p.lower() for p in re.split(r"(\d+)", s)]


def read_table(ws, key_header):
    """Rows of a sheet as dicts keyed by the header text in row 1.

    Rows whose `key_header` cell is blank are skipped. Columns are found by
    header name, so inserting or moving columns in Excel does not break this.
    """
    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        return [], []
    headers = [clean(h) for h in header]
    if key_header not in headers:
        raise KeyError(f"column '{key_header}' not found on the {ws.title} tab")
    out = []
    for r in rows:
        rec = {}
        for i, h in enumerate(headers):
            if h is None:
                continue
            rec[h] = clean(r[i]) if i < len(r) else None
        if rec.get(key_header) is None:
            continue
        out.append(rec)
    return out, [h for h in headers if h is not None]


def read_all_rows(ws):
    """Every non-empty row of a sheet as a list of cleaned values."""
    out = []
    for r in ws.iter_rows(values_only=True):
        vals = [clean(v) for v in r]
        while vals and vals[-1] is None:
            vals.pop()
        if vals:
            out.append(vals)
    return out


def read_bytes_with_retry(path):
    """Read the whole file into memory, retrying if it is mid-write or locked."""
    last = None
    for attempt in range(READ_RETRIES):
        try:
            with open(path, "rb") as f:
                data = f.read()
            # Opening as a zip proves the file is complete (xlsx is a zip).
            zipfile.ZipFile(io.BytesIO(data)).testzip()
            return data
        except (OSError, zipfile.BadZipFile) as e:
            last = e
            if attempt < READ_RETRIES - 1:
                time.sleep(RETRY_DELAY_S)
    raise last


# --------------------------------------------------------------------------
# Derived fields
# --------------------------------------------------------------------------

def cpu_short(cpu):
    """'Intel(R) Core(TM) i5-6500 CPU @ 3.20GHz' -> 'i5-6500'."""
    if not cpu:
        return None
    s = re.sub(r"@.*$", "", cpu)
    s = re.sub(r"\((R|TM|r|tm)\)", "", s)
    s = re.sub(r"\b(Intel|Core|CPU|Processor)\b", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s or cpu


RAM_GEN_PATTERNS = [
    (re.compile(r"DDR5|PC5", re.I), "DDR5"),
    (re.compile(r"DDR4|PC4", re.I), "DDR4"),
    (re.compile(r"DDR3L|PC3L", re.I), "DDR3L"),
    (re.compile(r"DDR3|PC3", re.I), "DDR3"),
    (re.compile(r"DDR2|PC2", re.I), "DDR2"),
]


def ram_group(spec):
    """'DIMM 8GB PC3L-12800U 2Rx8' -> 'DIMM 8GB DDR3L'; ECC flagged."""
    if not spec:
        return "Unknown"
    form = "SODIMM" if re.search(r"SODIMM", spec, re.I) else (
        "DIMM" if re.search(r"DIMM", spec, re.I) else "")
    size = re.search(r"(\d+)\s*GB", spec, re.I)
    gen = next((g for p, g in RAM_GEN_PATTERNS if p.search(spec)), "")
    ecc = " ECC" if re.search(r"\bECC\b|PC3L?-\d+R\b|PC4-\d+R\b", spec, re.I) else ""
    parts = [form, f"{size.group(1)}GB" if size else "", gen + ecc]
    label = " ".join(p for p in parts if p).strip()
    return label or spec


# --------------------------------------------------------------------------
# Snapshot
# --------------------------------------------------------------------------

def scan_files_by_asset(raw_dir):
    """{asset_id: [file names]} for every .txt in the scan folder."""
    out = defaultdict(list)
    if not raw_dir or not os.path.isdir(raw_dir):
        return out
    for name in sorted(os.listdir(raw_dir)):
        if not name.lower().endswith(".txt"):
            continue
        asset = name.split("_", 1)[0]
        out[asset].append(name)
    return out


def find_label_value(rows, label):
    """Value in the cell right of the first cell whose text equals `label`."""
    for r in rows:
        for i, v in enumerate(r):
            if isinstance(v, str) and v.strip().lower() == label.lower():
                return r[i + 1] if i + 1 < len(r) else None
    return None


def build_snapshot(path, raw_dir=None, baseline="default"):
    data = read_bytes_with_retry(path)
    # Two views of the same bytes: saved values, and the formulas themselves.
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    wbf = openpyxl.load_workbook(io.BytesIO(data), data_only=False, read_only=True)
    try:
        snap = _build(wb, wbf, path, raw_dir)
        if baseline == "default":
            baseline = health_mod.load_baseline()
        try:
            report = health_mod.check(wbf, data, baseline)
        except Exception as e:  # never let the health check take the dashboard down
            report = {"available": False, "error": f"{type(e).__name__}: {e}"}
        snap["health"] = report
        if report.get("problems"):
            n = len(report["problems"])
            snap["attention"].append({
                "level": "serious" if any(p["level"] == "serious" for p in report["problems"]) else "warning",
                "kind": "Workbook health", "asset_id": None,
                "text": f"{n} problem{'s' if n != 1 else ''} with formulas, drop-downs or shading rules. "
                        "See Workbook health on the Data tab."})
        return snap
    finally:
        wb.close()
        wbf.close()


def missing_formulas(wbf, sheet, key_header, formula_headers):
    """{key: [headers]} for data rows where an expected formula cell is not a formula."""
    if sheet not in wbf.sheetnames:
        return {}
    rows = wbf[sheet].iter_rows(values_only=True)
    header = [clean(h) for h in next(rows, [])]
    if key_header not in header:
        return {}
    ki = header.index(key_header)
    cols = [(h, header.index(h)) for h in formula_headers if h in header]
    out = {}
    for r in rows:
        key = clean(r[ki]) if ki < len(r) else None
        if key is None:
            continue
        miss = [h for h, i in cols
                if not (i < len(r) and isinstance(r[i], str) and r[i].startswith("="))]
        if miss:
            out[key] = miss
    return out


def _same(a, b):
    for k, v in a.items():
        w = b.get(k)
        if isinstance(v, (int, float)) and isinstance(w, (int, float)):
            if abs(v - w) > 0.005:
                return False
        elif (v or None) != (w or None):
            return False
    return True


def compute_sales(rows, known_ids):
    """Recompute the Sales tab's grey columns (Q-X) the way its formulas do.

    Used when Excel's saved results are missing, and to spot stale ones.
    """
    by_order = defaultdict(list)
    for r in rows:
        if r.get("Order ID"):
            by_order[r["Order ID"]].append(r)
    id_count = Counter(str(r["Asset ID"]).upper() for r in rows)

    def nz(v):
        n = num(v)
        return n if n is not None else 0

    out = []
    for r in rows:
        oid = r.get("Order ID")
        grp = by_order.get(oid, []) if oid else []
        first = grp[0] if grp else r
        # Q Price used
        if num(r.get("Sale price")) is not None:
            price = num(r.get("Sale price"))
        elif not oid or not any(num(g.get("Order total")) is not None for g in grp):
            price = None
        else:
            blanks = sum(1 for g in grp if num(g.get("Sale price")) is None)
            total = sum(nz(g.get("Order total")) for g in grp)
            typed = sum(nz(g.get("Sale price")) for g in grp)
            price = (total - typed) / blanks if blanks else None
        # R / S Fees and shipping used
        n = len(grp) or 1
        fees = num(r.get("Fees")) if num(r.get("Fees")) is not None else (
            0 if not oid else sum(nz(g.get("Order fees")) for g in grp) / n)
        ship = num(r.get("Shipping cost")) if num(r.get("Shipping cost")) is not None else (
            0 if not oid else sum(nz(g.get("Order shipping")) for g in grp) / n)
        net = price - fees - ship if price is not None else None
        # U Date used / V Platform used (fall back to the order's first row)
        date = r.get("Sale date") or (first.get("Sale date") if oid else None)
        platform = r.get("Platform") or (first.get("Platform") if oid else None)
        # X Check
        def cnt(col):
            return sum(1 for g in grp if g.get(col) is not None)
        if id_count[str(r["Asset ID"]).upper()] > 1:
            check = "DUPLICATE"
        elif str(r["Asset ID"]).upper() not in known_ids:
            check = "ID NOT FOUND"
        elif not oid and (nz(r.get("Order total")) + nz(r.get("Order fees")) + nz(r.get("Order shipping"))) > 0:
            check = "ORDER ID MISSING"
        elif oid and (cnt("Order total") > 1 or cnt("Order fees") > 1 or cnt("Order shipping") > 1):
            check = "ORDER TOTAL TWICE"
        elif price is None:
            check = "NO PRICE"
        elif not date:
            check = "NO DATE"
        elif oid and (nz(r.get("Fees")) + nz(r.get("Shipping cost"))) > 0 and \
                sum(nz(g.get("Order fees")) + nz(g.get("Order shipping")) for g in grp) > 0:
            check = "FEES TWICE"
        else:
            check = "OK"
            if price < 0:
                check = "PRICE MISMATCH"
            elif oid and cnt("Order total") > 0:
                prices = []
                for g in grp:
                    p = num(g.get("Sale price"))
                    if p is None:
                        blanks = sum(1 for x in grp if num(x.get("Sale price")) is None)
                        p = (sum(nz(x.get("Order total")) for x in grp)
                             - sum(nz(x.get("Sale price")) for x in grp)) / blanks
                    prices.append(p)
                if abs(sum(prices) - sum(nz(g.get("Order total")) for g in grp)) > 0.005:
                    check = "PRICE MISMATCH"
        aid = str(r["Asset ID"])
        out.append({
            "price": round(price, 2) if price is not None else None,
            "fees": round(fees, 2), "shipping": round(ship, 2),
            "net": round(net, 2) if net is not None else None,
            "date": date, "platform": platform, "check": check,
            "machines_in_order": len(grp) if oid else None,
            "type_guess": aid.split("-", 1)[0] if "-" in aid else None,
            "first": first,
        })
    return out


LISTING_STATUSES = ["Not listed", "Needs listed", "Listed", "Pending", "Removed"]
ACTIVE_LISTING = {"Listed", "Pending"}   # count toward the potential
STATUS_ORDER = ["Sold", "Pending", "Listed", "Needs listed", "Not listed", "Removed"]


def _uid(v):
    """Matching key for an asset ID or serial: text, trimmed, upper case
    (Excel's COUNTIF/MATCH ignore case)."""
    if v is None:
        return None
    s = str(v).strip()
    return s.upper() or None


def safe_url(v):
    """Only http(s) links are shown as links."""
    if isinstance(v, str) and re.match(r"^https?://", v.strip(), re.I):
        return v.strip()
    return None


def build_listings(listing_rows, computers, items, sold_upper, sale_by_upper,
                   sold_net, lot_price):
    """Combine the Listings tab with every machine and item.

    Each machine or item appears once. A unit on the Sales tab is Sold, whatever
    its Listings row says, and is counted only at its actual sale net. Unsold
    units count at their asking price only when their status is Listed or
    Pending and a price is entered.
    """
    attention = []
    # First row per ID wins; later duplicates are flagged and ignored.
    by_id = OrderedDict()
    for row_no, r in listing_rows:
        key = _uid(r.get("Asset ID"))
        if key in by_id:
            attention.append({"level": "serious", "kind": "Listing on two rows",
                              "asset_id": str(r.get("Asset ID")),
                              "text": f"Listings tab: this ID is on row {by_id[key][0]} and row {row_no}. "
                                      f"Only row {by_id[key][0]} is used. Delete the extra row."})
            continue
        by_id[key] = (row_no, r)

    units, used = [], set()

    def add_unit(kind, uid_list, base):
        listing, lkey = None, None
        for k in uid_list:
            if k and k in by_id:
                listing, lkey = by_id[k][1], k
                break
        if lkey:
            used.add(lkey)
        sold_key = next((k for k in uid_list if k and k in sold_upper), None)
        sale = sale_by_upper.get(sold_key) if sold_key else None
        raw_status = (listing or {}).get("Listing status")
        raw_status = str(raw_status).strip() if raw_status is not None else None
        canon = next((s for s in LISTING_STATUSES if raw_status and s.lower() == raw_status.lower()), None)
        if sold_key:
            status = "Sold"
        elif not raw_status:
            status = "Not listed"
        else:
            status = canon or raw_status
        asking = num((listing or {}).get("Asking price"))
        counts = (not sold_key and status in ACTIVE_LISTING and asking is not None and asking >= 0)
        rec = dict(base, kind=kind, status=status, listing_status=raw_status,
                   has_listing_row=listing is not None,
                   listing_row=by_id[lkey][0] if lkey else None,
                   asking=asking,
                   date_listed=(listing or {}).get("Date listed"),
                   platform=(listing or {}).get("Platform"),
                   url=safe_url((listing or {}).get("Listing URL")),
                   listing_notes=(listing or {}).get("Notes"),
                   sale_net=sale["net"] if sale else None,
                   sale_price=sale["price"] if sale else None,
                   counts=counts)
        units.append(rec)
        aid = base["id"] or "(no ID)"
        if not sold_key and raw_status and not canon:
            attention.append({"level": "warning", "kind": "Unknown listing status", "asset_id": base["id"],
                              "text": f"Listings tab: status '{raw_status}' is not one of "
                                      f"{', '.join(LISTING_STATUSES)}. Not counted in the potential."})
        if not sold_key and status in ACTIVE_LISTING and asking is None:
            attention.append({"level": "warning", "kind": "No asking price", "asset_id": base["id"],
                              "text": f"Listings tab: {aid} is {status} but has no asking price, "
                                      "so it is not counted in the potential."})
        if not sold_key and asking is not None and asking < 0:
            attention.append({"level": "warning", "kind": "Negative asking price", "asset_id": base["id"],
                              "text": "Listings tab: asking price is below zero. Not counted."})
        return rec

    for c in computers:
        rec = add_unit("Computer", [_uid(c["asset_id"])],
                       {"id": c["asset_id"], "type": c.get("type"), "model": c.get("model"),
                        "detail": " · ".join(str(x) for x in [c.get("cpu_short"),
                                             f"{c['ram_gb']} GB" if c.get("ram_gb") is not None else None] if x),
                        "in_computers": True})
        c["listing"] = {k: rec[k] for k in ("status", "listing_status", "asking", "date_listed",
                                            "platform", "url", "listing_notes", "has_listing_row", "counts")}
    for it in items:
        aid, serial = it.get("Asset ID"), it.get("Serial")
        rec = add_unit("Hardware", [_uid(aid), _uid(serial)],
                 {"id": str(aid) if aid is not None else (str(serial) if serial is not None else None),
                  "serial": str(serial) if serial is not None else None,
                  "type": it.get("Type"), "model": it.get("Model"), "detail": it.get("Key spec"),
                  "in_computers": False})
        it["_listing"] = {k: rec[k] for k in ("status", "listing_status", "asking", "date_listed",
                                              "platform", "url", "listing_notes", "has_listing_row", "counts")}

    # Listings rows that match no machine or item.
    for key, (row_no, r) in by_id.items():
        if key in used:
            continue
        rid = str(r.get("Asset ID"))
        rec = add_unit("Not in inventory", [key], {"id": rid, "type": None, "model": None,
                                                    "detail": None, "in_computers": False})
        rec["counts"] = False
        attention.append({"level": "warning", "kind": "Listing ID not found", "asset_id": rid,
                          "text": f"Listings tab row {row_no}: this ID is not on the Computers, Master "
                                  "or Items tab. Check for a typo. Not counted in the potential."})

    counted = [u for u in units if u["counts"]]
    asking_total = round(sum(u["asking"] for u in counted), 2)
    potential = round(sold_net + asking_total, 2)
    # Rows that match nothing are shown in the table but are not real units.
    real = [u for u in units if u["kind"] != "Not in inventory"]
    counts = Counter(u["status"] for u in real)
    status_counts = [{"label": s, "count": counts[s]} for s in STATUS_ORDER if counts.get(s)]
    status_counts += [{"label": s, "count": n} for s, n in counts.items() if s not in STATUS_ORDER]
    active = [u for u in real if u["status"] in ACTIVE_LISTING]
    # The same counts split by kind: machines (Computers/Master tabs) and
    # hardware (Items tab).
    kind_counts = {k: Counter(u["status"] for u in real if u["kind"] == k)
                   for k in ("Computer", "Hardware")}
    for row in status_counts:
        row["machines"] = kind_counts["Computer"].get(row["label"], 0)
        row["other"] = kind_counts["Hardware"].get(row["label"], 0)

    def split(status):
        return {"machines": kind_counts["Computer"].get(status, 0),
                "other": kind_counts["Hardware"].get(status, 0)}
    summary = {
        "present": True,
        "rows": len(listing_rows),
        "status_counts": status_counts,
        "active": len(active),
        "listed": counts.get("Listed", 0),
        "pending": counts.get("Pending", 0),
        "needs_listed": counts.get("Needs listed", 0),
        "by_status": {s: split(s) for s in STATUS_ORDER},
        "active_split": {"machines": sum(1 for u in active if u["kind"] == "Computer"),
                         "other": sum(1 for u in active if u["kind"] == "Hardware")},
        "totals_split": {"machines": sum(kind_counts["Computer"].values()),
                         "other": sum(kind_counts["Hardware"].values())},
        "active_priced": len(counted),
        "active_unpriced": sum(1 for u in active if u["asking"] is None),
        "asking_total": asking_total,
        "sold_net": round(sold_net, 2),
        "net_sales_potential": potential,
        "profit_potential": round(potential - lot_price, 2) if lot_price is not None else None,
    }
    return units, summary, attention


COMPUTER_FORMULA_COLS = ["win11_override", "manual_notes", "win11_effective", "ram_mb_effective", "sold"]
SALES_FORMULA_COLS = ["Type", "Model", "Price used", "Fees used", "Shipping used", "Net",
                      "Date used", "Platform used", "Machines in order", "Check"]


def _sheet(wb, name, warnings_out, required=False):
    if name in wb.sheetnames:
        return wb[name]
    msg = f"The '{name}' tab is missing from the workbook."
    if required:
        raise KeyError(msg)
    warnings_out.append(msg)
    return None


def _build(wb, wbf, path, raw_dir):
    warn = []
    st = os.stat(path)

    # ---- Raw sheets --------------------------------------------------------
    ws = _sheet(wb, "Computers", warn, required=True)
    computers, computer_cols = read_table(ws, "asset_id")

    ws = _sheet(wb, "Sales", warn)
    sales_raw, sales_cols = read_table(ws, "Asset ID") if ws else ([], [])

    ws = _sheet(wb, "Master", warn)
    master, master_cols = read_table(ws, "Asset ID") if ws else ([], [])

    ws = _sheet(wb, "Items", warn)
    items, item_cols = ([], [])
    if ws:
        # RAM rows have no Asset ID, so keep any row with something in it.
        rows = ws.iter_rows(values_only=True)
        header = [clean(h) for h in next(rows, [])]
        item_cols = [h for h in header if h]
        for row_no, r in enumerate(rows, start=2):
            rec = {h: clean(r[i]) if i < len(r) else None
                   for i, h in enumerate(header) if h}
            if any(v is not None for v in rec.values()):
                rec["_row"] = row_no
                items.append(rec)

    ws = _sheet(wb, "Bulk", warn)
    bulk, bulk_cols = ([], [])
    if ws:
        rows = ws.iter_rows(values_only=True)
        header = [clean(h) for h in next(rows, [])]
        bulk_cols = [h for h in header if h]
        for r in rows:
            rec = {h: clean(r[i]) if i < len(r) else None
                   for i, h in enumerate(header) if h}
            if any(v is not None for v in rec.values()):
                bulk.append(rec)

    ws = _sheet(wb, "Selling Dashboard", warn)
    sd_rows = read_all_rows(ws) if ws else []
    lot_price = num(find_label_value(sd_rows, "Lot purchase price (enter)"))
    if lot_price is None:
        warn.append("Lot purchase price not found on the Selling Dashboard tab "
                    "(cell next to 'Lot purchase price (enter)'). Profit and "
                    "recovery figures are hidden.")

    ws = _sheet(wb, "Summary", warn)
    summary_rows = read_all_rows(ws) if ws else []

    ws = _sheet(wb, "Lists", warn)
    lists = OrderedDict()
    if ws:
        rows = read_all_rows(ws)
        if rows:
            header = rows[0]
            for i, h in enumerate(header):
                if h:
                    lists[h] = [r[i] for r in rows[1:] if i < len(r) and r[i] is not None]

    # Listings is optional: without it the dashboard works as before.
    listing_rows, listing_cols = [], []
    has_listings = "Listings" in wb.sheetnames
    if has_listings:
        rows = wb["Listings"].iter_rows(values_only=True)
        header = [clean(h) for h in next(rows, [])]
        listing_cols = [h for h in header if h]
        if "Asset ID" not in header:
            warn.append("The Listings tab has no 'Asset ID' column in row 1, so listings are ignored.")
        else:
            for row_no, r in enumerate(rows, start=2):
                rec = {h: clean(r[i]) if i < len(r) else None for i, h in enumerate(header) if h}
                if rec.get("Asset ID") is not None:
                    listing_rows.append((row_no, rec))

    # ---- Indexes -----------------------------------------------------------
    # Excel's MATCH is case-insensitive, so the lookups here are too.
    master_by_id = {}
    for m in master:
        master_by_id.setdefault(str(m["Asset ID"]).upper(), m)
    computer_ids = {c["asset_id"] for c in computers}
    scans = scan_files_by_asset(raw_dir)

    # ---- Formula integrity -------------------------------------------------
    comp_missing = missing_formulas(wbf, "Computers", "asset_id", COMPUTER_FORMULA_COLS)
    sales_missing = missing_formulas(wbf, "Sales", "Asset ID", SALES_FORMULA_COLS)
    stale_sales = []
    uncached_sales = [s["Asset ID"] for s in sales_raw
                      if s.get("Check") is None and s["Asset ID"] not in sales_missing]
    if uncached_sales:
        warn.append(
            f"{len(uncached_sales)} Sales row(s) have no saved formula results "
            "(the file was saved by something that did not recalculate). The "
            "dashboard filled them in itself; open the workbook in Excel and save "
            "it to be sure the numbers match.")

    # ---- Sales -------------------------------------------------------------
    item_ids = {str(i.get(k)).upper() for i in items for k in ("Asset ID", "Serial") if i.get(k)}
    known_ids = ({str(x).upper() for x in computer_ids} | set(master_by_id) | item_ids)
    computed = compute_sales(sales_raw, known_ids)
    sales = []
    sold_ids = set()
    for s, calc in zip(sales_raw, computed):
        has_saved = s.get("Check") is not None
        if has_saved:
            saved = {"price": num(s.get("Price used")), "fees": num(s.get("Fees used")),
                     "shipping": num(s.get("Shipping used")), "net": num(s.get("Net")),
                     "date": s.get("Date used"), "platform": s.get("Platform used"),
                     "check": s.get("Check")}
            if not _same(saved, calc):
                stale_sales.append(s["Asset ID"])
            use = saved
        else:
            use = calc
        first = calc["first"]
        rec = {
            "asset_id": s["Asset ID"],
            "date": use["date"],
            "order_id": s.get("Order ID"),
            "type": s.get("Type") or calc["type_guess"],
            "model": s.get("Model"),
            "price": use["price"],
            "fees": use["fees"] or 0,
            "shipping": use["shipping"] or 0,
            "net": use["net"],
            "platform": use["platform"],
            "payment": s.get("Payment method") or first.get("Payment method"),
            "buyer": s.get("Buyer (optional)") or first.get("Buyer (optional)"),
            "sold_with": s.get("Sold with"),
            "notes": s.get("Notes"),
            "machines_in_order": calc["machines_in_order"],
            "check": use["check"],
            "in_computers": s["Asset ID"] in computer_ids,
        }
        sales.append(rec)
        sold_ids.add(s["Asset ID"])
    sales.sort(key=lambda r: (r["date"] or "", r["order_id"] or "", natural_key(r["asset_id"])),
               reverse=True)
    sale_by_id = {s["asset_id"]: s for s in sales}

    # ---- Computers ---------------------------------------------------------
    out_comp = []
    sold_ids_upper = {s.upper() for s in sold_ids}
    stale = []
    for c in computers:
        aid = c["asset_id"]
        m = master_by_id.get(str(aid).upper())
        # Columns U-Y, computed with the same rules as the workbook formulas
        # (README section A, step 10). This keeps the dashboard right even
        # when a paste has wiped a formula in Excel.
        override = clean(str((m or {}).get("Win11 override") or "")) if m else None
        notes = clean(str((m or {}).get("Notes") or "")) if m else None
        win11 = override or c.get("win11")
        ram_mb_raw = num(c.get("ram_mb"))
        ram_mb = 0 if notes and "real ram = none" in notes.lower() else ram_mb_raw
        sold = str(aid).upper() in sold_ids_upper

        # Where the workbook has a working formula, its saved value should agree.
        if aid not in comp_missing:
            saved = {"win11_effective": c.get("win11_effective"), "ram_mb_effective": num(c.get("ram_mb_effective")),
                     "sold": c.get("sold") == "SOLD", "manual_notes": c.get("manual_notes")}
            ours = {"win11_effective": win11, "ram_mb_effective": ram_mb, "sold": sold, "manual_notes": notes}
            if saved["ram_mb_effective"] is not None and saved != ours:
                stale.append(aid)

        c = dict(c, win11_override=override, manual_notes=notes, win11_effective=win11,
                 ram_mb_effective=ram_mb, sold="SOLD" if sold else None)
        tpm = c.get("tpm") or ""
        rec = dict(c)
        rec.update({
            "cpu_short": cpu_short(c.get("cpu")),
            "ram_gb": round(ram_mb / 1024, 1) if ram_mb is not None else None,
            "win11_eff": win11,
            "tpm_visible": tpm.upper().startswith("TPM VISIBLE"),
            "has_drive": bool(c.get("drives")),
            "is_sold": bool(sold),
            "master": m,
            "sale": sale_by_id.get(aid),
            "scan_files": scans.get(aid, []),
            "source": "scan",
        })
        if isinstance(rec.get("ram_gb"), float) and rec["ram_gb"].is_integer():
            rec["ram_gb"] = int(rec["ram_gb"])
        out_comp.append(rec)

    # Machines that exist only on the Master tab (never scanned).
    for m in master:
        aid = m["Asset ID"]
        if aid in computer_ids:
            continue
        out_comp.append({
            "asset_id": aid, "type": m.get("Type"), "service_tag": m.get("Service Tag"),
            "model": m.get("Model"), "status": m.get("Status code") or "NOT SCANNED",
            "manual_notes": m.get("Notes"), "win11_eff": m.get("Win11 override"),
            "cpu_short": None, "ram_gb": None, "tpm_visible": None, "has_drive": None,
            "is_sold": str(aid).upper() in sold_ids_upper, "master": m, "sale": sale_by_id.get(aid),
            "scan_files": scans.get(aid, []), "source": "master-only",
        })
    out_comp.sort(key=lambda r: natural_key(r["asset_id"]))

    # ---- KPIs --------------------------------------------------------------
    priced = [s for s in sales if s["price"] is not None]
    gross = sum(s["price"] for s in priced)
    fees = sum(s["fees"] or 0 for s in sales)
    ship = sum(s["shipping"] or 0 for s in sales)
    net = sum(s["net"] for s in sales if s["net"] is not None)
    unsold = [c for c in out_comp if not c["is_sold"]]
    kpis = {
        "units_sold": len(sales),
        "orders": len({s["order_id"] or f"_{s['asset_id']}" for s in sales}),
        "gross": round(gross, 2),
        "fees": round(fees, 2),
        "shipping": round(ship, 2),
        "net": round(net, 2),
        "avg_price": round(gross / len(priced), 2) if priced else None,
        "avg_net": round(net / len(priced), 2) if priced else None,
        "lot_price": lot_price,
        "recovered": round(net / lot_price, 4) if lot_price else None,
        "profit": round(net - lot_price, 2) if lot_price is not None else None,
        "machines_total": len(out_comp),
        "machines_scanned": len(computers),
        "machines_unsold": len(unsold),
        "first_sale": min((s["date"] for s in sales if s["date"]), default=None),
        "last_sale": max((s["date"] for s in sales if s["date"]), default=None),
    }
    if lot_price and kpis["avg_net"] and net < lot_price:
        remaining = lot_price - net
        kpis["break_even_remaining"] = round(remaining, 2)
        kpis["break_even_units"] = -(-remaining // kpis["avg_net"])  # ceil
    else:
        kpis["break_even_remaining"] = 0 if lot_price and net >= lot_price else None
        kpis["break_even_units"] = 0 if kpis["break_even_remaining"] == 0 else None

    # ---- Listings and potential ---------------------------------------------
    if has_listings:
        sale_by_upper = {}
        for s in sales:
            sale_by_upper.setdefault(_uid(s["asset_id"]), s)
        listing_units, listings, listing_attention = build_listings(
            listing_rows, out_comp, items, set(sale_by_upper), sale_by_upper, net, lot_price)
    else:
        listing_units, listing_attention = [], []
        listings = {"present": False}
    # ---- Hardware (Items tab) -----------------------------------------------
    sale_by_upper_all = {}
    for s in sales:
        sale_by_upper_all.setdefault(_uid(s["asset_id"]), s)
    hardware = []
    for it in items:
        aid, serial = it.get("Asset ID"), it.get("Serial")
        sale = sale_by_upper_all.get(_uid(aid)) or sale_by_upper_all.get(_uid(serial))
        hid = str(aid) if aid is not None else (str(serial) if serial is not None else None)
        hardware.append({
            "id": hid, "asset_id": aid, "serial": serial,
            "type": it.get("Type"), "make": it.get("Make"), "model": it.get("Model"),
            "key_spec": it.get("Key spec"), "condition": it.get("Condition"), "notes": it.get("Notes"),
            "ram_group": ram_group(it.get("Key spec")) if (it.get("Type") or "").upper() == "RAM" else None,
            "row": it.get("_row"), "is_sold": sale is not None, "sale": sale,
            "listing": it.get("_listing"),
            "extra": {k: v for k, v in it.items() if not k.startswith("_") and k not in
                      ("Asset ID", "Serial", "Type", "Make", "Model", "Key spec", "Condition", "Notes")},
        })
    hw_ids = Counter(h["id"].upper() for h in hardware if h["id"])
    for h in hardware:
        h["id_unique"] = bool(h["id"]) and hw_ids[h["id"].upper()] == 1

    kpis["net_sales_potential"] = listings.get("net_sales_potential")
    kpis["profit_potential"] = listings.get("profit_potential")

    # By type (machines) -- scanned + master-only.
    by_type = OrderedDict()
    for c in out_comp:
        t = c.get("type") or "?"
        b = by_type.setdefault(t, {"type": t, "total": 0, "sold": 0, "unsold": 0, "net": 0.0, "priced": 0})
        b["total"] += 1
        if c["is_sold"]:
            b["sold"] += 1
            s = c.get("sale")
            if s and s["net"] is not None:
                b["net"] += s["net"]
                b["priced"] += 1
        else:
            b["unsold"] += 1
    other_sales = [s for s in sales if s["asset_id"] not in {c["asset_id"] for c in out_comp}]
    if other_sales:
        b = by_type.setdefault("Other hardware", {"type": "Other hardware", "total": None,
                                                  "sold": 0, "unsold": None, "net": 0.0, "priced": 0})
        for s in other_sales:
            b["sold"] += 1
            if s["net"] is not None:
                b["net"] += s["net"]
                b["priced"] += 1
    for b in by_type.values():
        b["net"] = round(b["net"], 2)
        b["avg_net"] = round(b["net"] / b["priced"], 2) if b["priced"] else None

    def group_sum(key):
        g = OrderedDict()
        for s in sales:
            k = s.get(key) or "(blank)"
            e = g.setdefault(k, {"label": k, "count": 0, "gross": 0.0, "net": 0.0})
            e["count"] += 1
            e["gross"] += s["price"] or 0
            e["net"] += s["net"] or 0
        res = sorted(g.values(), key=lambda e: -e["net"])
        for e in res:
            e["gross"] = round(e["gross"], 2)
            e["net"] = round(e["net"], 2)
        return res

    by_day = OrderedDict()
    for s in sorted(sales, key=lambda s: s["date"] or ""):
        d = s["date"] or "(no date)"
        e = by_day.setdefault(d, {"date": d, "count": 0, "gross": 0.0, "net": 0.0})
        e["count"] += 1
        e["gross"] += s["price"] or 0
        e["net"] += s["net"] or 0
    cum = 0.0
    for e in by_day.values():
        cum += e["net"]
        e["gross"] = round(e["gross"], 2)
        e["net"] = round(e["net"], 2)
        e["cumulative_net"] = round(cum, 2)

    # ---- Needs attention ---------------------------------------------------
    attention = []
    for s in sales:
        if s["check"] != "OK":
            attention.append({"level": "serious", "kind": "Sale check",
                              "asset_id": s["asset_id"],
                              "text": f"Sales row check says {s['check']}."})
    for c in out_comp:
        if c["source"] == "master-only":
            attention.append({"level": "warning", "kind": "Not scanned",
                              "asset_id": c["asset_id"],
                              "text": "On the Master tab but has no scan row. "
                                      + (f"Note: {c['manual_notes']}" if c.get("manual_notes") else "")})
            continue
        if c.get("status") and c["status"] != "OK":
            prob = c.get("problems")
            attention.append({"level": "warning", "kind": f"Status {c['status']}",
                              "asset_id": c["asset_id"],
                              "text": f"Scan status is {c['status']}" + (f" ({prob})." if prob else ".")})
        note = c.get("manual_notes") or ""
        if re.search(r"scan row", note, re.I):
            attention.append({"level": "info", "kind": "Note mentions the scan row",
                              "asset_id": c["asset_id"],
                              "text": "Check whether this Master note still applies to the current "
                                      "scan row: " + note})
    for aid, cols in sorted(comp_missing.items(), key=lambda kv: natural_key(kv[0])):
        attention.append({"level": "warning", "kind": "Formula missing in Excel", "asset_id": aid,
                          "text": "Computers tab: no formula in " + ", ".join(cols) + ", so Excel leaves "
                                  "these blank for this row. Copy the formulas down from a row above. "
                                  "The dashboard computes these columns itself, so its values are "
                                  "correct either way."})
    for aid, cols in sales_missing.items():
        attention.append({"level": "serious", "kind": "Formula missing in Excel", "asset_id": aid,
                          "text": "Sales tab: no formula in " + ", ".join(cols) + ". Copy the formulas "
                                  "down from a row above so the Sales totals stay right."})
    for aid in stale_sales:
        attention.append({"level": "warning", "kind": "Saved values out of date", "asset_id": aid,
                          "text": "Sales tab: Excel's saved results for this row differ from a fresh "
                                  "calculation. Open the workbook in Excel and save it."})
    for aid in stale:
        attention.append({"level": "warning", "kind": "Saved values out of date", "asset_id": aid,
                          "text": "Excel's saved values for columns U-Y differ from the Master/Sales "
                                  "tabs. Open the workbook in Excel and save it."})
    for m in master:
        code = m.get("Status code")
        if code and code != "OK":
            attention.append({"level": "warning", "kind": f"Master status {code}",
                              "asset_id": m["Asset ID"], "text": m.get("Notes") or ""})

    attention.extend(listing_attention)

    # Cross-check against the workbook's own Selling Dashboard totals.
    checks = []
    for label, ours in [("Units sold", kpis["units_sold"]), ("Gross sales", kpis["gross"]),
                        ("Net", kpis["net"])]:
        theirs = num(find_label_value(sd_rows, label))
        if theirs is None:
            continue
        ok = abs(theirs - ours) < 0.005
        checks.append({"label": label, "dashboard": ours, "workbook": theirs, "ok": ok})
        if not ok:
            attention.append({"level": "serious", "kind": "Totals mismatch", "asset_id": None,
                              "text": f"{label}: dashboard {ours} vs Selling Dashboard tab {theirs}. "
                                      "Open the workbook in Excel and save it so formulas recalculate."})

    sheets = OrderedDict()
    for name in wb.sheetnames:
        sheets[name] = None
    sheets.update({"Computers": len(computers), "Sales": len(sales_raw), "Master": len(master),
                   "Items": len(items), "Bulk": len(bulk)})
    if has_listings:
        sheets["Listings"] = len(listing_rows)

    return {
        "meta": {
            "file": os.path.basename(path),
            "mtime": dt.datetime.fromtimestamp(st.st_mtime, dt.timezone.utc).isoformat(),
            "size": st.st_size,
            "read_at": dt.datetime.now(dt.timezone.utc).isoformat(),
            "raw_dir_found": bool(raw_dir and os.path.isdir(raw_dir)),
            "warnings": warn,
            "sheets": sheets,
            "checks": checks,
        },
        "kpis": kpis,
        "by_type": list(by_type.values()),
        "by_platform": group_sum("platform"),
        "by_payment": group_sum("payment"),
        "by_model_sold": group_sum("model"),
        "by_day": list(by_day.values()),
        "computers": out_comp,
        "sales": sales,
        "master": master,
        "master_cols": master_cols,
        "items": [{k: v for k, v in i.items() if not k.startswith("_")} for i in items],
        "hardware": hardware,
        "item_cols": item_cols,
        "ram_groups": [{"label": k, "count": v} for k, v in Counter(
            ram_group(i.get("Key spec")) for i in items if (i.get("Type") or "").upper() == "RAM"
        ).most_common()],
        "bulk": bulk,
        "bulk_cols": bulk_cols,
        "summary_rows": summary_rows,
        "lists": lists,
        "attention": attention,
        "listings": listings,
        "listing_units": listing_units,
        "listing_cols": listing_cols,
    }
