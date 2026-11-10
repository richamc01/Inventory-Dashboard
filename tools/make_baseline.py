"""Save the workbook's current formulas, drop-downs and shading rules as the
reference copy (dashboard/baseline.json) that the Data tab's health check uses.

Run it after a deliberate change to the workbook's formulas or drop-downs:
    python3 tools/make_baseline.py "/path/to/Hardware Inventory.xlsx"

Only formula text and drop-down/shading definitions are saved, never cell data.
"""
import io
import json
import os
import sys
import warnings

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from dashboard.health import BASELINE_PATH, make_baseline  # noqa: E402

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

LAST = 500
FILL_DOWN = (
    [("Computers", h, 2, LAST) for h in
     ["win11_override", "manual_notes", "win11_effective", "ram_mb_effective", "sold"]] +
    [("Sales", h, 2, LAST) for h in
     ["Type", "Model", "Price used", "Fees used", "Shipping used", "Net", "Date used",
      "Platform used", "Machines in order", "Check"]] +
    [("Listings", "Current status", 2, LAST)]
)
FIXED = ["Summary", "Selling Dashboard"]


def main(src, dst=BASELINE_PATH):
    data = open(src, "rb").read()
    wbf = openpyxl.load_workbook(io.BytesIO(data), read_only=True)
    base = make_baseline(wbf, data, FILL_DOWN, FIXED)
    with open(dst, "w", encoding="utf-8") as f:
        json.dump(base, f, indent=1, ensure_ascii=False)
        f.write("\n")
    n_fixed = sum(len(v) for v in base["fixed"].values())
    n_dv = sum(len(v) for v in base["data_validations"].values())
    n_cf = sum(len(v) for v in base["conditional_formats"].values())
    print(f"{dst}: {len(base['fill_down'])} fill-down columns, {n_fixed} one-off formulas, "
          f"{n_dv} drop-downs/entry rules, {n_cf} shading rules, {len(base['defined_names'])} named ranges")


if __name__ == "__main__":
    main(*sys.argv[1:])
