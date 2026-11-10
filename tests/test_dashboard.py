import os
import shutil
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from dashboard.app import create_app  # noqa: E402
from dashboard.workbook import build_snapshot, cpu_short, ram_group  # noqa: E402
from make_fixture import build  # noqa: E402


@pytest.fixture()
def fx(tmp_path):
    path = build(str(tmp_path))
    return path, str(tmp_path / "inventory-raw")


@pytest.fixture()
def snap(fx):
    return build_snapshot(*fx)


def comp(snap, aid):
    return next(c for c in snap["computers"] if c["asset_id"] == aid)


# ---- workbook logic --------------------------------------------------------

def test_cpu_short():
    assert cpu_short("Intel(R) Core(TM) i5-6500 CPU @ 3.20GHz") == "i5-6500"
    assert cpu_short("Intel(R) Xeon(R) CPU E5-2430 v2 @ 2.50GHz") == "Xeon E5-2430 v2"
    assert cpu_short(None) is None


def test_ram_group():
    assert ram_group("DIMM 8GB PC3L-12800U 2Rx8") == "DIMM 8GB DDR3L"
    assert ram_group("DIMM 16GB PC3L-12800R ECC 2Rx4") == "DIMM 16GB DDR3L ECC"
    assert ram_group("SODIMM 4GB DDR4-2400") == "SODIMM 4GB DDR4"


def test_counts(snap):
    k = snap["kpis"]
    assert k["machines_scanned"] == 5
    assert k["machines_total"] == 6          # + AIO-004 from Master only
    assert k["units_sold"] == 3
    assert k["machines_unsold"] == 3          # DT-003, DT-010, AIO-004
    assert k["lot_price"] == 500


def test_sales_math_from_formulas(snap):
    s = {x["asset_id"]: x for x in snap["sales"]}
    # Single sale with a fee.
    assert s["DT-001"]["price"] == 100 and s["DT-001"]["fees"] == 10 and s["DT-001"]["net"] == 90
    # Order total 100 split evenly over two machines; second row inherits date/platform.
    assert s["DT-002"]["price"] == 50 and s["SV-001"]["price"] == 50
    assert s["SV-001"]["date"] == "2026-01-07"
    assert s["SV-001"]["platform"] == "Facebook Marketplace"
    assert s["SV-001"]["payment"] == "Cash"
    assert s["SV-001"]["machines_in_order"] == 2
    assert all(x["check"] == "OK" for x in snap["sales"])
    k = snap["kpis"]
    assert k["gross"] == 200 and k["net"] == 190 and k["orders"] == 2
    assert k["recovered"] == pytest.approx(0.38)
    assert k["profit"] == -310
    assert k["break_even_units"] == 5        # 310 / 63.33 -> 5


def test_cross_check_matches(snap):
    assert snap["meta"]["checks"] and all(c["ok"] for c in snap["meta"]["checks"])


def test_master_lookups_computed(snap):
    assert comp(snap, "DT-002")["win11_eff"] == "YES-IF-TPM-ON"
    assert comp(snap, "DT-002")["manual_notes"] == "Fake note."
    # Formula wiped in Excel, but the dashboard still finds the Master note.
    assert comp(snap, "DT-003")["manual_notes"] == "Wiped-formula row note."
    # "real RAM = none" -> effective RAM 0.
    assert comp(snap, "DT-010")["ram_gb"] == 0
    assert comp(snap, "DT-001")["is_sold"] and not comp(snap, "DT-003")["is_sold"]


def test_master_only_machine(snap):
    a = comp(snap, "AIO-004")
    assert a["source"] == "master-only" and a["status"] == "NOT SCANNED"


def test_attention(snap):
    kinds = {(a["asset_id"], a["kind"]) for a in snap["attention"]}
    assert ("DT-003", "Formula missing in Excel") in kinds
    assert ("AIO-004", "Not scanned") in kinds
    assert ("DT-010", "Status CHECK") in kinds


def test_scan_files(snap):
    assert comp(snap, "DT-010")["scan_files"] == ["DT-010_TAG0010.txt", "DT-010_TAG0010_120000.txt"]


def test_items_and_bulk(snap):
    assert len(snap["items"]) == 3 and len(snap["bulk"]) == 1
    assert {"label": "DIMM 16GB DDR3L ECC", "count": 1} in snap["ram_groups"]


def test_sale_problem_flagged(tmp_path):
    import openpyxl
    path = build(str(tmp_path))
    wb = openpyxl.load_workbook(path)
    wb["Sales"].append([None, "DT-001", None, None, None, 5])   # duplicate + no date
    wb.save(path)
    s = build_snapshot(path, None)
    checks = [x["check"] for x in s["sales"] if x["asset_id"] == "DT-001"]
    assert checks == ["DUPLICATE", "DUPLICATE"]
    assert any(a["kind"] == "Sale check" for a in s["attention"])


def test_missing_lot_price_warns(tmp_path):
    import openpyxl
    path = build(str(tmp_path))
    wb = openpyxl.load_workbook(path)
    wb["Selling Dashboard"]["B3"] = None
    wb.save(path)
    s = build_snapshot(path, None)
    assert s["kpis"]["lot_price"] is None and s["kpis"]["profit"] is None
    assert any("Lot purchase price" in w for w in s["meta"]["warnings"])


# ---- HTTP ------------------------------------------------------------------

@pytest.fixture()
def client(fx):
    app = create_app(fx[0], fx[1])
    return app.test_client()


def test_pages(client):
    assert client.get("/").status_code == 200
    for f in ("app.js", "style.css"):
        assert client.get(f"/static/{f}").status_code == 200
    assert client.get("/healthz").data == b"ok\n"


def test_api_data(client):
    r = client.get("/api/data")
    assert r.status_code == 200 and r.headers["Cache-Control"] == "no-store"
    j = r.get_json()
    assert j["kpis"]["units_sold"] == 3 and j["meta"]["error"] is None


def test_raw_file(client):
    r = client.get("/api/raw/DT-001")
    assert r.status_code == 200 and b"asset_id=DT-001" in r.data
    r = client.get("/api/raw/DT-010?file=DT-010_TAG0010_120000.txt")
    assert b"rescan=yes" in r.data


@pytest.mark.parametrize("url", [
    "/api/raw/DT-001?file=../Hardware%20Inventory.xlsx",
    "/api/raw/DT-001?file=/etc/passwd",
    "/api/raw/AIO-004",
    "/api/raw/NOPE",
    "/static/../app.py",
])
def test_no_path_escape(client, url):
    assert client.get(url).status_code == 404


def test_missing_workbook_503(tmp_path):
    c = create_app(str(tmp_path / "missing.xlsx")).test_client()
    r = c.get("/api/data")
    assert r.status_code == 503 and "Cannot open the workbook" in r.get_json()["error"]
    assert c.get("/api/status").get_json()["ok"] is False
    assert c.get("/healthz").status_code == 200


def test_bad_update_keeps_last_good(fx, monkeypatch):
    import dashboard.workbook as wbmod
    monkeypatch.setattr(wbmod, "RETRY_DELAY_S", 0)
    app = create_app(*fx)
    c = app.test_client()
    assert c.get("/api/data").status_code == 200
    # Simulate a half-written file.
    with open(fx[0], "r+b") as f:
        f.truncate(1000)
    r = c.get("/api/data")
    assert r.status_code == 200
    j = r.get_json()
    assert j["kpis"]["units_sold"] == 3 and "Could not read" in j["meta"]["error"]
    # Restore -> recovers.
    build(os.path.dirname(fx[0]))
    j = c.get("/api/data").get_json()
    assert j["meta"]["error"] is None


def test_reloads_when_file_changes(fx):
    import openpyxl
    c = create_app(*fx).test_client()
    assert c.get("/api/data").get_json()["kpis"]["units_sold"] == 3
    wb = openpyxl.load_workbook(fx[0])
    wb["Sales"].append(["2026-01-09", "DT-003", None, None, None, 40])
    wb.save(fx[0])
    os.utime(fx[0], (1, 2_000_000_000))   # force a different mtime
    j = c.get("/api/data").get_json()
    assert j["kpis"]["units_sold"] == 4


REAL = os.environ.get("REAL_WORKBOOK")


@pytest.mark.skipif(not REAL, reason="set REAL_WORKBOOK to test against the real file")
def test_real_workbook():
    s = build_snapshot(REAL, os.path.join(os.path.dirname(REAL), "inventory-raw"))
    assert s["meta"]["checks"] and all(c["ok"] for c in s["meta"]["checks"]), s["meta"]["checks"]
    assert not any(a["kind"] == "Saved values out of date" for a in s["attention"])
    assert s["kpis"]["machines_scanned"] > 0
    assert s["health"]["available"] and s["health"]["ok"], s["health"]["problems"]
    for c in s["computers"]:
        if c["source"] == "scan":
            assert c["scan_files"], c["asset_id"]


# ---- Listings tab and potential ---------------------------------------------

@pytest.fixture()
def lsnap(tmp_path):
    path = build(str(tmp_path), listings=True)
    return build_snapshot(path, str(tmp_path / "inventory-raw"))


def unit(snap, uid):
    return next(u for u in snap["listing_units"] if u["id"] == uid)


def test_no_listings_tab(snap):
    assert snap["listings"] == {"present": False}
    assert snap["listing_units"] == []
    assert snap["kpis"]["net_sales_potential"] is None
    assert snap["kpis"]["profit_potential"] is None
    assert not any("isting" in a["kind"] for a in snap["attention"])


def test_potential_math(lsnap):
    L = lsnap["listings"]
    # Sold net: DT-001 90 + DT-002 50 + SV-001 50.
    assert L["sold_net"] == 190
    # Counted asking prices: DT-003 80 (Listed), DT-010 60 (Pending), MON-001 40 (Listed item).
    assert L["asking_total"] == 180
    assert L["net_sales_potential"] == 370
    assert L["profit_potential"] == 370 - 500
    assert lsnap["kpis"]["net_sales_potential"] == 370
    assert lsnap["kpis"]["profit_potential"] == -130
    assert L["active_priced"] == 3 and L["active_unpriced"] == 1   # SER-RAM-2 has no price


def test_sold_unit_counted_once_at_sale_net(lsnap):
    u = unit(lsnap, "DT-001")
    assert u["status"] == "Sold" and u["asking"] == 150 and u["counts"] is False and u["sale_net"] == 90


def test_statuses(lsnap):
    assert unit(lsnap, "DT-003")["status"] == "Listed"
    assert unit(lsnap, "DT-010")["status"] == "Pending"          # 'pending' typed in lower case
    assert unit(lsnap, "SER-RAM-1")["status"] == "Needs listed" and not unit(lsnap, "SER-RAM-1")["counts"]
    assert unit(lsnap, "AIO-004")["status"] == "Removed" and not unit(lsnap, "AIO-004")["counts"]
    assert unit(lsnap, "DT-002")["status"] == "Sold" and not unit(lsnap, "DT-002")["has_listing_row"]
    assert unit(lsnap, "MON-001")["kind"] == "Hardware"
    counts = {s["label"]: s["count"] for s in lsnap["listings"]["status_counts"]}
    # 6 machines + 4 items; XX-999 matches nothing and is not a unit.
    assert counts == {"Sold": 3, "Pending": 1, "Listed": 3, "Needs listed": 1, "Removed": 1, "On hold": 1}
    assert sum(counts.values()) == 10
    L = lsnap["listings"]
    # 5 machines + AIO-004 (Master only) = 6 machines; 4 hardware items.
    assert L["totals_split"] == {"machines": 6, "other": 4}
    assert L["by_status"]["Sold"] == {"machines": 3, "other": 0}
    assert L["by_status"]["Listed"] == {"machines": 1, "other": 2}     # DT-003; MON-001, SER-RAM-2
    assert L["by_status"]["Pending"] == {"machines": 1, "other": 0}
    assert L["by_status"]["Removed"] == {"machines": 1, "other": 0}
    assert L["active_split"] == {"machines": 2, "other": 2}
    rows = {r["label"]: r for r in L["status_counts"]}
    assert all(r["machines"] + r["other"] == r["count"] for r in rows.values())


def test_listing_problems_flagged(lsnap):
    kinds = {(a["kind"], a["asset_id"]) for a in lsnap["attention"]}
    assert ("Listing on two rows", "dt-003") in kinds
    assert ("Listing ID not found", "XX-999") in kinds
    assert ("No asking price", "SER-RAM-2") in kinds
    assert ("Unknown listing status", "SER-RAM-3") in kinds
    # The duplicate row's 999 is ignored; the first row's 80 is used.
    assert unit(lsnap, "DT-003")["asking"] == 80
    assert not unit(lsnap, "XX-999")["counts"] and not unit(lsnap, "SER-RAM-3")["counts"]
    # A sold machine still on the Listings tab is normal, not a problem.
    assert not any(a["asset_id"] == "DT-001" and "isting" in a["kind"] for a in lsnap["attention"])


def test_listing_urls_and_drawer(lsnap):
    assert unit(lsnap, "DT-003")["url"] is None                     # javascript: link dropped
    assert unit(lsnap, "MON-001")["url"] == "https://example.com/listing/1"
    c = comp(lsnap, "DT-003")
    assert c["listing"]["status"] == "Listed" and c["listing"]["asking"] == 80 and c["listing"]["counts"]
    assert comp(lsnap, "DT-002")["listing"]["has_listing_row"] is False


def test_listings_api(tmp_path):
    path = build(str(tmp_path), listings=True)
    app = create_app(path, str(tmp_path / "inventory-raw"))
    body = app.test_client().get("/api/data").get_json()
    assert body["listings"]["net_sales_potential"] == 370
    assert body["meta"]["sheets"]["Listings"] == 10


# ---- Hardware inventory -----------------------------------------------------

def test_hardware_list(snap):
    hw = {h["id"]: h for h in snap["hardware"]}
    assert set(hw) == {"SER-RAM-1", "SER-RAM-2", "MON-001"}       # RAM uses its serial as ID
    assert hw["MON-001"]["serial"] == "SER-MON-1" and hw["MON-001"]["make"] == "Dell"
    assert hw["SER-RAM-2"]["ram_group"] == "DIMM 16GB DDR3L ECC"
    assert hw["MON-001"]["ram_group"] is None
    assert all(h["id_unique"] and not h["is_sold"] for h in hw.values())
    assert hw["SER-RAM-1"]["row"] == 2
    assert "_row" not in snap["items"][0]


def test_hardware_sold_by_serial(tmp_path):
    import openpyxl
    path = build(str(tmp_path))
    wb = openpyxl.load_workbook(path)
    wb["Sales"].append([None, "SER-MON-1", None, None, None, 40, None, None, None, None, None, "eBay"])
    wb["Sales"]["A5"] = wb["Sales"]["A2"].value
    wb.save(path)
    s = build_snapshot(path, None, baseline=None)
    mon = next(h for h in s["hardware"] if h["id"] == "MON-001")
    assert mon["is_sold"] and mon["sale"]["price"] == 40


def test_hardware_listing_attached(lsnap):
    mon = next(h for h in lsnap["hardware"] if h["id"] == "MON-001")
    assert mon["listing"]["status"] == "Listed" and mon["listing"]["asking"] == 40


# ---- Workbook health ---------------------------------------------------------

from dashboard import health  # noqa: E402


def test_shift_rows():
    f = '=IF($A2="","",IF(COUNTIF(Sales!$B$2:$B$500,$A2)>0,"SOLD A2",""))'
    assert health.shift_rows(f, 3) == '=IF($A5="","",IF(COUNTIF(Sales!$B$2:$B$500,$A5)>0,"SOLD A2",""))'
    assert health.shift_rows("=SUM(C2:C9)+LOG10(4)+B$2", 1) == "=SUM(C3:C10)+LOG10(4)+B$2"
    assert health.ranges([2, 3, 4, 9, 11, 12]) == "2-4, 9, 11-12"


def _health_fixture(tmp_path):
    """Fixture with formulas filled down, a drop-down and a shading rule, plus its baseline."""
    import io
    import openpyxl
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    path = build(str(tmp_path), listings=True)
    wb = openpyxl.load_workbook(path)
    cw = wb["Computers"]
    for r in range(2, 11):
        cw[f"Y{r}"] = f'=IF($A{r}="","",IF(COUNTIF(Sales!$B$2:$B$500,$A{r})>0,"SOLD",""))'
    sd = wb["Selling Dashboard"]
    sd["B9"] = "=SUM(Sales!$T$2:$T$500)"
    dv = DataValidation(type="list", formula1="Lists!$B$2:$B$2", allow_blank=True)
    dv.add("L2:L500")
    wb["Sales"].add_data_validation(dv)
    cw.conditional_formatting.add("A2:Y500", FormulaRule(formula=['$Y2="SOLD"']))
    wb.save(path)
    data = open(path, "rb").read()
    wbf = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
    base = health.make_baseline(wbf, data, [("Computers", "sold", 2, 10)], ["Selling Dashboard"])
    return path, base


def test_health_clean(tmp_path):
    path, base = _health_fixture(tmp_path)
    s = build_snapshot(path, None, baseline=base)
    h = s["health"]
    assert h["available"] and h["ok"], h["problems"]
    assert h["checked"] == 9 + 1 + 1 + 1
    assert not any(a["kind"] == "Workbook health" for a in s["attention"])


def test_health_finds_problems(tmp_path):
    import openpyxl
    path, base = _health_fixture(tmp_path)
    wb = openpyxl.load_workbook(path)
    cw = wb["Computers"]
    cw["Y4"] = None                    # wiped
    cw["Y5"] = None
    cw["Y8"] = '=IF($A8="","","SOLD")'  # edited
    wb["Selling Dashboard"]["B9"] = 190  # formula typed over
    wb["Sales"].data_validations.dataValidation = []
    cw.conditional_formatting = type(cw.conditional_formatting)()
    wb.save(path)
    s = build_snapshot(path, None, baseline=base)
    texts = [p["text"] for p in s["health"]["problems"]]
    assert any("no formula in row(s) 4-5" in t for t in texts)
    assert any("differs" in t and "row(s) 8" in t for t in texts)
    assert any("No formula in B9" in t for t in texts)
    assert any("drop-down from Lists!$B$2:$B$2 is missing" in t for t in texts)
    assert any("Shading rule" in t and "missing" in t for t in texts)
    assert any(a["kind"] == "Workbook health" for a in s["attention"])


def test_health_moved_column(tmp_path):
    import openpyxl
    path, base = _health_fixture(tmp_path)
    wb = openpyxl.load_workbook(path)
    wb["Computers"]["Y1"] = "renamed"
    wb.save(path)
    probs = build_snapshot(path, None, baseline=base)["health"]["problems"]
    assert any("header is 'renamed'" in p["text"] for p in probs)


def test_health_without_baseline(fx):
    s = build_snapshot(*fx, baseline=None)
    assert s["health"] == {"available": False}


def test_bundled_baseline_is_formula_only():
    import json
    base = json.load(open(health.BASELINE_PATH))
    assert set(base) == {"version", "fill_down", "fixed", "data_validations", "conditional_formats",
                         "defined_names"}
    for cells in base["fixed"].values():
        assert all(v.startswith("=") for v in cells.values())


# ---- Type codes and instructions ----------------------------------------------

def test_type_codes_from_lists(tmp_path):
    import openpyxl
    path = build(str(tmp_path))
    wb = openpyxl.load_workbook(path)
    ls = wb["Lists"]
    ls["J1"], ls["K1"], ls["L1"] = "Type code", "Meaning", "Used for"
    ls["J2"], ls["K2"], ls["L2"] = "DT", "Desktops", "Computers"
    ls["J3"], ls["K3"], ls["L3"] = "RAM", "Memory modules", "Hardware"
    wb.save(path)
    s = build_snapshot(path, None, baseline=None)
    assert s["type_codes"] == [{"code": "DT", "meaning": "Desktops", "used_for": "Computers"},
                               {"code": "RAM", "meaning": "Memory modules", "used_for": "Hardware"}]
    # The type-code columns are not shown as drop-down lists.
    assert "Type code" not in s["lists"] and "Meaning" not in s["lists"]
    assert s["lists"]["Type"] == ["MON"]


def test_no_type_codes(snap):
    assert snap["type_codes"] == []


def test_hardware_ram_parts(snap):
    ram = {h["id"]: h["ram"] for h in snap["hardware"] if h["ram"]}
    assert ram["SER-RAM-1"] == {"form": "DIMM", "size_gb": 4, "gen": "DDR3L", "ecc": False}
    assert ram["SER-RAM-2"] == {"form": "DIMM", "size_gb": 16, "gen": "DDR3L", "ecc": True}


def test_instructions_page(fx):
    import re
    app = create_app(*fx)
    r = app.test_client().get("/static/instructions.html")
    assert r.status_code == 200
    html = r.get_data(as_text=True)
    assert "<!--TYPE_CODES-->" not in html
    ids = set(re.findall(r'id="([^"]+)"', html))
    jumps = set(re.findall(r'data-jump="([^"]+)"', html))
    assert jumps and jumps <= ids, jumps - ids
    assert html.count("<div") == html.count("</div>")


# ---- Final statuses -----------------------------------------------------------

def test_final_status_not_flagged(tmp_path):
    import openpyxl
    path = build(str(tmp_path))
    wb = openpyxl.load_workbook(path)
    ms = wb["Master"]
    for r in range(2, ms.max_row + 1):
        if ms.cell(r, 1).value == "AIO-004":
            ms.cell(r, 5).value = "DEAD"          # master-only machine, final
        if ms.cell(r, 1).value == "DT-002":
            ms.cell(r, 5).value = "BIOS-PW"       # scanned machine, final
        if ms.cell(r, 1).value == "DT-003":
            ms.cell(r, 5).value = "NO-POST"       # not final: still flagged
    wb.save(path)
    s = build_snapshot(path, None, baseline=None)
    flagged = {(a["kind"], a["asset_id"]) for a in s["attention"]}
    assert ("Not scanned", "AIO-004") not in flagged
    assert ("Master status DEAD", "AIO-004") not in flagged
    assert ("Master status BIOS-PW", "DT-002") not in flagged
    assert ("Master status NO-POST", "DT-003") in flagged
    # Still in inventory, with its final status.
    aio = next(c for c in s["computers"] if c["asset_id"] == "AIO-004")
    assert aio["final_status"] == "DEAD" and aio["status"] == "DEAD" and not aio["is_sold"]
    assert next(c for c in s["computers"] if c["asset_id"] == "DT-002")["final_status"] == "BIOS-PW"
    assert s["kpis"]["machines_total"] == 6
