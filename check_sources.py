"""Cheap source-change probe for the poll-triggered refresh.

Asks Graph for every source file's lastModifiedDateTime (metadata only — no
downloads) and compares against sources_meta.json committed by the last
build. Prints changed=true/false to GITHUB_OUTPUT so the workflow rebuilds
only when something actually moved: fresh Paylocity hours land minutes
after the 6 AM Eastern relay, and same-day debrief edits are picked up
within the poll interval instead of waiting for tomorrow.

Fail-open: a missing baseline or a metadata error counts as changed, so a
broken probe degrades to "build anyway" (the build's own retries and loud
failures take it from there), never to a silently stale page.
"""
import json
import os
import sys
from pathlib import Path

import requests

from fetch_sources import DRIVE_ID, get_token
from sources import SOURCE_FILES as FILES
from sources import SOURCE_FOLDERS as FOLDERS

BASELINE = Path(__file__).parent / "sources_meta.json"


def main():
    old = {}
    if BASELINE.exists():
        try:
            old = json.loads(BASELINE.read_text())
        except ValueError:
            print("baseline unreadable — treating everything as changed")
    else:
        print("no committed baseline (first poll run) — building")

    changed = []
    hdrs = {"Authorization": f"Bearer {get_token()}"}
    for path in FILES:
        name = path.rsplit("/", 1)[-1]
        url = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
               f"/root:/{requests.utils.quote(path)}"
               f"?$select=lastModifiedDateTime")
        try:
            r = requests.get(url, headers=hdrs, timeout=30)
            r.raise_for_status()
            stamp = r.json().get("lastModifiedDateTime")
        except requests.RequestException as e:
            print(f"  {name}: metadata error ({type(e).__name__}) — "
                  f"counting as changed")
            changed.append(name)
            continue
        if old.get(name) != stamp:
            print(f"  CHANGED {name}: {old.get(name)} -> {stamp}")
            changed.append(name)

    # Mirrored folders are polled the same way, by listing their children.
    # Without this an edit to a Private/MRO forecast board did not trigger a
    # rebuild at all: the page only caught up when some OTHER source moved,
    # typically the next morning's Paylocity file. Reported 2026-10-23 (Sam)
    # on AFW EMB, whose jobs had sat unpublished for a day. An added or
    # removed file counts too — that is how a new location arrives.
    for folder in FOLDERS:
        base = folder.rsplit("/", 1)[-1]
        url = (f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}"
               f"/root:/{requests.utils.quote(folder)}:/children"
               f"?$select=name,lastModifiedDateTime&$top=200")
        try:
            r = requests.get(url, headers=hdrs, timeout=30)
            r.raise_for_status()
            kids = {f"{base}/{c['name']}": c.get("lastModifiedDateTime")
                    for c in r.json().get("value", [])
                    if c["name"].lower().endswith(".json")}
        except requests.RequestException as e:
            print(f"  {folder}: listing error ({type(e).__name__}) — "
                  f"counting as changed")
            changed.append(folder)
            continue
        seen = {k for k in old if k.startswith(base + "/")}
        for name, stamp in sorted(kids.items()):
            if old.get(name) != stamp:
                print(f"  CHANGED {name}: {old.get(name)} -> {stamp}")
                changed.append(name)
        for gone in sorted(seen - set(kids)):
            print(f"  REMOVED {gone}")
            changed.append(gone)

    verdict = "true" if (changed or not old) else "false"
    print(f"{len(changed)} source(s) changed -> changed={verdict}")
    out = os.environ.get("GITHUB_OUTPUT")
    if out:
        with open(out, "a") as f:
            f.write(f"changed={verdict}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
