"""One-off (2026-08-31): set the Private and MRO Pulse Locations price
matrices (TUS MHI + SLN 1V) to the QB-billed prices — matrix x 1.035, the
uplift every 2026 TUS invoice carries (Sam). Cell-level PATCHes through the
Graph WORKBOOK API per the live-workbook rules; run via workflow_dispatch
with the Report Automation app creds. Delete after the run.
"""
import json
import sys

import requests

from fetch_sources import DRIVE_ID, get_token

PATH = "Pulse%20Sheets/Private%20and%20MRO%20Pulse%20Locations.xlsx"
BASE = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
        f"/root:/{PATH}:/workbook")

# (range address, values matrix) — F2:H9 minus the blank Brightwork cells
PATCHES = [
    ("F2:H3", [[1732.59, 1878.53, 1961.33], [822.83, 1029.83, 1133.33]]),
    ("F4",    [[2887.65]]),
    ("F5:H9", [[1707.75, 1707.75, 1707.75],
               [258.75, 362.25, 439.88],
               [15275.57, 16827.03, 19056.42],
               [2277, 2794.5, 2794.5],
               [879.75, 1009.13, 1009.13]]),
]
SHEETS = ["TUS MHI", "SLN 1V"]


def main():
    hdrs = {"Authorization": f"Bearer {get_token()}",
            "Content-Type": "application/json"}
    for sheet in SHEETS:
        s = requests.utils.quote(sheet)
        for addr, values in PATCHES:
            r = requests.patch(
                f"{BASE}/worksheets('{s}')/range(address='{addr}')",
                headers=hdrs, json={"values": values}, timeout=60)
            r.raise_for_status()
            print(f"PATCHED {sheet} {addr}")
        chk = requests.get(
            f"{BASE}/worksheets('{s}')/range(address='E2:H9')?$select=values",
            headers=hdrs, timeout=60)
        chk.raise_for_status()
        print(f"VERIFY {sheet}:")
        for row in chk.json()["values"]:
            print("  ", row)
    print("DONE")
    return 0


if __name__ == "__main__":
    sys.exit(main())
