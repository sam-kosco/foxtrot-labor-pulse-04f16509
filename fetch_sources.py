"""Downloads the pulse source files from SharePoint via Microsoft Graph into
./sources (the same files the Live Commercial Pulse Sheet queries).

Credentials (GitHub Secrets / env): TENANT_ID, CLIENT_ID, CLIENT_SECRET —
the Foxtrot Report Automation app, same as the compliance trackers.
"""
import os
import sys
import time
from pathlib import Path

import requests

from sources import SOURCE_FILES as FILES
from sources import SOURCE_FOLDERS as FOLDERS

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

# Data Hub document library (drive-relative paths, no "Shared Documents" prefix)
DRIVE_ID = "b!_bzXaIx86kOufgJN3ih-BaDIDthKYuxJkJtLi1Bm5irGjCEnK-VHSpBRRm3_SDKU"

OUT = Path(__file__).parent / "sources"
# The stamps of the bytes CURRENTLY IN ./sources — written beside them by
# whichever run produced them, and restored with them from the runner cache.
# A file whose Graph lastModifiedDateTime still matches has not moved, so
# there is nothing to download. The poll (check_sources.py) has always
# trusted this exact comparison to decide whether to build at all; this
# makes the FETCH trust it too.
#
# Deliberately NOT the committed sources_meta.json at the repo root: that
# describes the last BUILD, while the cache may have come from a different
# run. Reading the stamps that travelled WITH the bytes is the only version
# of this that cannot skip a download over a stale file.
BASELINE = OUT / "sources_meta.json"


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
    import json
    OUT.mkdir(exist_ok=True)
    token = get_token()
    print("Token acquired.")
    failed, meta = [], {}
    hdrs = {"Authorization": f"Bearer {token}"}
    try:
        prev = json.loads(BASELINE.read_text())
    except (OSError, ValueError):
        prev = {}        # no baseline (or unreadable) = fetch everything
        print("no committed baseline — downloading every source")
    kept = 0

    def unchanged(key, stamp):
        """Is the cached copy already this exact version?"""
        dest = OUT / key
        return bool(stamp) and prev.get(key) == stamp \
            and dest.exists() and dest.stat().st_size > 0
    for path in FILES:
        base = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
                f"/root:/{requests.utils.quote(path)}")
        name = path.rsplit("/", 1)[-1]
        m = requests.get(base + "?$select=lastModifiedDateTime", headers=hdrs, timeout=30)
        if m.status_code == 200:
            meta[name] = m.json().get("lastModifiedDateTime")
        if unchanged(name, meta.get(name)):
            kept += 1
            continue
        ok = False
        for attempt in range(1, 4):
            try:
                r = requests.get(base + ":/content", headers=hdrs, timeout=180)
                if r.status_code == 200:
                    (OUT / name).write_bytes(r.content)
                    print(f"  ok  {name} ({len(r.content) // 1024} KB, "
                          f"modified {meta.get(name)})")
                    ok = True
                    break
                print(f"  retry {name}: HTTP {r.status_code} (attempt {attempt})",
                      file=sys.stderr)
            except requests.RequestException as e:
                print(f"  retry {name}: {type(e).__name__} (attempt {attempt})",
                      file=sys.stderr)
            time.sleep(5 * attempt)
        if not ok:
            failed.append(name)
            print(f"  FAIL {name}: giving up after 3 attempts", file=sys.stderr)
    # Folders are mirrored whole: their contents change as locations are
    # added, so a fixed file list would go stale without anyone noticing.
    for folder in FOLDERS:
        dest = OUT / folder.rsplit("/", 1)[-1]
        dest.mkdir(exist_ok=True)
        url = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
               f"/root:/{requests.utils.quote(folder)}:/children"
               f"?$select=name,lastModifiedDateTime&$top=200")
        r = requests.get(url, headers=hdrs, timeout=60)
        if r.status_code != 200:
            failed.append(f"{folder} (listing HTTP {r.status_code})")
            continue
        kids = [c for c in r.json().get("value", [])
                if c["name"].lower().endswith(".json")]
        for c in kids:
            key = f"{dest.name}/{c['name']}"
            stamp = c.get("lastModifiedDateTime")
            if unchanged(key, stamp):
                meta[key] = stamp
                kept += 1
                continue
            cu = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/root:/"
                  f"{requests.utils.quote(folder + '/' + c['name'])}:/content")
            cr = requests.get(cu, headers=hdrs, timeout=120)
            if cr.status_code == 200:
                (dest / c["name"]).write_bytes(cr.content)
                meta[key] = stamp
            else:
                failed.append(f"{folder}/{c['name']} (HTTP {cr.status_code})")
        # A child the folder no longer holds must not linger in the cache:
        # the build reads the whole directory, so a deleted location would
        # keep publishing from a stale file.
        live = {c["name"] for c in kids}
        for old_file in dest.glob("*.json"):
            if old_file.name not in live:
                old_file.unlink()
                print(f"  gone {folder}/{old_file.name} — removed from cache")
        print(f"  ok  {folder}: {len(kids)} file(s)")
        if not kids:
            failed.append(f"{folder} (empty — Private/MRO budgets would be 0)")

    if kept:
        print(f"  {kept} source(s) unchanged — served from cache, not downloaded")
    (OUT / "sources_meta.json").write_text(json.dumps(meta, indent=1))
    if failed:
        sys.exit(f"aborting: {len(failed)} downloads failed {failed}")


if __name__ == "__main__":
    main()
