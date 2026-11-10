"""Build a small fake workbook with the same layout as Hardware Inventory.xlsx.

No real inventory data. Formula columns hold formulas with no saved results
(openpyxl cannot calculate), so the dashboard's own calculations are exercised.

Usage: python tests/make_fixture.py OUT_DIR
Creates OUT_DIR/Hardware Inventory.xlsx and OUT_DIR/inventory-raw/*.txt
"""

import datetime as dt
import os
import sys

import openpyxl

COMP_HEADERS = ["asset_id", "type", "service_tag", "vendor", "model", "cpu", "cpu_gen", "ram_mb", "dimms",
                "drives", "mac_addresses", "bios_version", "tpm", "bios_tpm_settings", "win11",
                "low_ram_under_4gb", "no_drive_64gb_plus", "status", "problems", "detail_file",
                "win11_override", "manual_notes", "win11_effective", "ram_mb_effective", "sold"]

SALES_HEADERS = ["Sale date", "Asset ID", "Order ID", "Type", "Model", "Sale price", "Fees", "Shipping cost",
                 "Order total", "Order fees", "Order shipping", "Platform", "Payment method", "Buyer (optional)",
                 "Sold with", "Notes", "Price used", "Fees used", "Shipping used", "Net", "Date used",
                 "Platform used", "Machines in order", "Check"]

MACHINES = [
    # asset, type, tag, model, cpu, gen, ram, drives, tpm, win11, status, problems
    ("DT-001", "DT", "TAG0001", "OptiPlex 5070", "Intel(R) Core(TM) i5-9500 CPU @ 3.00GHz", 9, 16384, None,
     "TPM VISIBLE, version_major=2", "YES", "OK", None),
    ("DT-002", "DT", "TAG0002", "OptiPlex 3040", "Intel(R) Core(TM) i5-6500 CPU @ 3.20GHz", 6, 8192,
     "FAKE SSD 240 GB", "NO TPM VISIBLE", "NO-CPU", "OK", None),
    ("DT-003", "DT", "TAG0003", "OptiPlex 3040", "Intel(R) Core(TM) i5-6500 CPU @ 3.20GHz", 6, 8192, None,
     "NO TPM VISIBLE", "NO-CPU", "OK", None),
    ("DT-010", "DT", "TAG0010", "OptiPlex 7040", "Intel(R) Core(TM) i5-6500 CPU @ 3.20GHz", 6, 4096, None,
     "TPM VISIBLE, version_major=2", "NO-CPU", "CHECK", "no-drive"),
    ("SV-001", "SV", "TAG0101", "PowerEdge R320", "Intel(R) Xeon(R) CPU E5-2430 v2 @ 2.50GHz", "unknown",
     24576, "PERC H710 3598 GB", "NO TPM VISIBLE", "N/A", "OK", None),
]


def formula_row(r):
    return [
        f'=IF($A{r}="","",IFERROR(INDEX(Master!$G$2:$G$499,MATCH($A{r},Master!$A$2:$A$499,0))&"",""))',
        f'=IF($A{r}="","",IFERROR(INDEX(Master!$H$2:$H$499,MATCH($A{r},Master!$A$2:$A$499,0))&"",""))',
        f'=IF($A{r}="","",IF($U{r}<>"",$U{r},$O{r}))',
        f'=IF($A{r}="","",IF(ISNUMBER(SEARCH("real RAM = none",$V{r})),0,$H{r}))',
        f'=IF($A{r}="","",IF(COUNTIF(Sales!$B$2:$B$500,$A{r})>0,"SOLD",""))',
    ]


def build(out_dir):
    os.makedirs(os.path.join(out_dir, "inventory-raw"), exist_ok=True)
    wb = openpyxl.Workbook()
    wb.active.title = "README"
    wb["README"]["A1"] = "Fake test workbook"

    ws = wb.create_sheet("Summary")
    ws.append(["Inventory summary"])
    ws.append([])
    ws.append(["Computers by type"])
    ws.append(["DT", 4])

    sd = wb.create_sheet("Selling Dashboard")
    sd.append(["Selling Dashboard"])
    sd.append([])
    sd.append(["Lot purchase price (enter)", 500, "Yellow cell"])
    sd.append(["Totals"])
    sd.append(["Units sold", 3])
    sd.append(["Gross sales", 200])
    sd.append(["Net", 190])

    it = wb.create_sheet("Items")
    it.append(["Asset ID", "Serial", "Type", "Make", "Model", "Key spec", "Condition", "Notes"])
    it.append([None, "SER-RAM-1", "RAM", "Samsung", "M378B5173EB0-YK0", "DIMM 4GB DDR3L-1600", "Untested", None])
    it.append([None, "SER-RAM-2", "RAM", "SK hynix", "HMT42GR7BFR4A-PB", "DIMM 16GB PC3L-12800R ECC 2Rx4", "Untested", None])
    it.append(["MON-001", "SER-MON-1", "MON", "Dell", "P2419H", "24in", "Working", None])

    bk = wb.create_sheet("Bulk")
    bk.append(["Bin ID", "Item type", "Spec", "Quantity", "Condition"])
    bk.append(["BIN-001", "Cable", "DisplayPort 6 ft", 12, "Working"])

    cw = wb.create_sheet("Computers")
    cw.append(COMP_HEADERS)
    for i, m in enumerate(MACHINES):
        r = i + 2
        aid, typ, tag, model, cpu, gen, ram, drives, tpm, win11, status, problems = m
        row = [aid, typ, tag, "Dell Inc.", model, cpu, gen, ram, 1, drives, "eth0=00:00:00:00:00:0" + str(i),
               "1.0.0", tpm, None, win11, "N", "N" if drives else "Y", status, problems, f"{aid}_{tag}.txt"]
        if aid == "DT-003":
            # Simulate a paste that wiped the U and V formulas.
            row += [None, None] + formula_row(r)[2:]
        else:
            row += formula_row(r)
        cw.append(row)
        with open(os.path.join(out_dir, "inventory-raw", f"{aid}_{tag}.txt"), "w") as f:
            f.write(f"== META\nasset_id={aid}\n\n== WIN11 SUMMARY\ncpu_model={cpu}\n")
    # A second scan of DT-010.
    with open(os.path.join(out_dir, "inventory-raw", "DT-010_TAG0010_120000.txt"), "w") as f:
        f.write("== META\nasset_id=DT-010\nrescan=yes\n")

    ms = wb.create_sheet("Master")
    ms.append(["Asset ID", "Service Tag", "Model", "Type", "Status code", "Date", "Win11 override", "Notes"])
    ms.append(["AIO-004", None, None, "AIO", None, None, None, "Will not boot."])
    ms.append(["DT-002", None, None, "DT", None, None, "YES-IF-TPM-ON", "Fake note."])
    ms.append(["DT-003", None, None, "DT", None, None, None, "Wiped-formula row note."])
    ms.append(["DT-010", None, None, "DT", None, None, None, "TEST DIMM USED, real RAM = none"])

    sa = wb.create_sheet("Sales")
    sa.append(SALES_HEADERS)
    sales = [
        # date, asset, order, price, fees, ship, total, ofees, oship, platform, pay, buyer
        (dt.datetime(2026, 1, 5), "DT-001", None, 100, 10, None, None, None, None, "eBay", "Platform payout", "A"),
        (dt.datetime(2026, 1, 7), "DT-002", "ORD-1", None, None, None, 100, None, None, "Facebook Marketplace", "Cash", "B"),
        (None, "SV-001", "ORD-1", None, None, None, None, None, None, None, None, None),
    ]
    for i, s in enumerate(sales):
        r = i + 2
        date, aid, oid, price, fees, ship, total, ofees, oship, plat, pay, buyer = s
        sa.append([date, aid, oid, f'=IF(B{r}="","","x")', f'=IF(B{r}="","","x")', price, fees, ship, total,
                   ofees, oship, plat, pay, buyer, None, None] + [f'=IF(B{r}="","",0)'] * 8)

    ls = wb.create_sheet("Lists")
    ls.append(["Type", "Sale platform"])
    ls.append(["MON", "eBay"])

    path = os.path.join(out_dir, "Hardware Inventory.xlsx")
    wb.save(path)
    return path


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "fixture"))
