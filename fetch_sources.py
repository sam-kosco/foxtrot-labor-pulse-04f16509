"""Downloads the pulse source files from SharePoint via Microsoft Graph into
./sources (the same files the Live Commercial Pulse Sheet queries).

Credentials (GitHub Secrets / env): TENANT_ID, CLIENT_ID, CLIENT_SECRET —
the Foxtrot Report Automation app, same as the compliance trackers.
"""
import os
import sys
from pathlib import Path

import requests

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

# Data Hub document library (drive-relative paths, no "Shared Documents" prefix)
DRIVE_ID = "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU"

FILES = [
    "Paylocity Reports/This Years Hours.csv",
    "Paylocity Reports/Basic Employee Info.csv",
    "Power Flows/Debriefs/Envoy Debriefs.xlsx",
    "Power Flows/Debriefs/GoJet Debriefs.xlsx",
    "Power Flows/Debriefs/Mesa Debriefs.xlsx",
    "Power Flows/Debriefs/PSA Debriefs.xlsx",
    "Power Flows/Debriefs/Breeze Debriefs.xlsx",
    "Power Flows/Debriefs/Ultra Debriefs.xlsx",
    "Power Flows/Debriefs/Frontier Debriefs.xlsx",
    "Power Flows/Debriefs/AA Debriefs.csv",
    "Power Flows/Debriefs/APU Wash.xlsx",
    "Power BI Data Sources/Location Management.csv",
    "Pulse Sheets/Service Budgets.xlsx",
]

OUT = Path(__file__).parent / "sources"


def get_token():
    resp = requests.post(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def main():
    OUT.mkdir(exist_ok=True)
    token = get_token()
    print("Token acquired.")
    failed = []
    for path in FILES:
        url = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
               f"/root:/{requests.utils.quote(path)}:/content")
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120)
        name = path.rsplit("/", 1)[-1]
        if r.status_code == 200:
            (OUT / name).write_bytes(r.content)
            print(f"  ok  {name} ({len(r.content) // 1024} KB)")
        else:
            failed.append((name, r.status_code))
            print(f"  FAIL {name}: HTTP {r.status_code}", file=sys.stderr)
    if failed:
        sys.exit(f"aborting: {len(failed)} downloads failed {failed}")


if __name__ == "__main__":
    main()
