"""Builds pulse_data.json from the same Data Hub sources the Live Commercial
Pulse Sheet queries — the workbook itself is never read.

Replicates: Power Query transforms (per-airline debrief normalization), the
sheet helper columns (punch-in minus 12h day attribution, labor-dist repair,
pay-type lookup), and the sheet formulas (COUNTIFS service counts, budgeted
= rate x count + fixed elapsed-day budgets, worked = hourly punches + 40/7
per active salaried head, weekly averages over elapsed days, variance text).
"""
import json
import os
import re
from calendar import monthrange
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
import pandas as pd

HERE = Path(__file__).parent
if os.environ.get("PULSE_DATA_DIR"):
    # CI mode: fetch_sources.py has downloaded everything into one flat dir
    DEBRIEFS = PAYLOCITY = LOCMGMT_DIR = BUDGETS_DIR = DEFINITIVE_DIR = \
        Path(os.environ["PULSE_DATA_DIR"])
else:
    DH = Path(r"C:\Users\samko\Foxtrot Aviation Services\Data Hub - Documents")
    DEBRIEFS = DH / "Power Flows" / "Debriefs"
    PAYLOCITY = DH / "Paylocity Reports"
    LOCMGMT_DIR = DH / "Power BI Data Sources"
    BUDGETS_DIR = DH / "Pulse Sheets"
    DEFINITIVE_DIR = DH / "Definitive Lists"
WINDOW_START = date(2026, 5, 1)  # same cutoff the workbook queries hardcode
EXCLUDED_TITLES = {"Director", "Regional Director", "Regional Manager II"}
TODAY = date.today()


def read_table(path, table_name):
    """Read a named Excel table into a DataFrame (matches what Power Query's
    Source{[Item=...,Kind="Table"]} sees)."""
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    for ws in wb.worksheets:
        if table_name in ws.tables:
            ref = ws.tables[table_name].ref
            rows = list(ws[ref])
            header = [c.value for c in rows[0]]
            data = [[c.value for c in r] for r in rows[1:]]
            wb.close()
            return pd.DataFrame(data, columns=header)
    wb.close()
    raise KeyError(f"table {table_name!r} not in {path.name}")


def yes(v):
    return 1 if isinstance(v, str) and v.strip().startswith("Yes") else 0


def not_no(v):
    return 0 if (isinstance(v, str) and v.strip() == "No") else (1 if v else 0)


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return pd.to_datetime(v).date() if v else None


# ---------------------------------------------------------------- debriefs

def load_envoy():
    main = read_table(DEBRIEFS / "Envoy Debriefs.xlsx", "Table1")
    main = pd.DataFrame({
        "date": main["Date"].map(to_date),
        "Location": main["Location"].astype(str).str.strip(),
        "Tail": main["Tail Number"],
        "IHC": main["IHC"].map(yes),
        "RRON": 0,
        "ED1": main["Exterior Detail #1 (ED1)"].map(yes),
        "ED2": main["Exterior Detail #2 (ED2)"].map(yes),
    })
    dfw = read_table(DEBRIEFS / "Envoy Debriefs.xlsx", "DFW_Debriefs")
    dfw = pd.DataFrame({
        "date": dfw["Date"].map(to_date),
        "Location": "DFW",
        "Tail": dfw["Tail"],
        "IHC": dfw["IHC"].map(yes),
        "RRON": dfw["RRON"].map(yes),
        "ED1": dfw["ED1"].map(yes),
        "ED2": dfw["ED2"].map(yes),
    })
    return pd.concat([main, dfw], ignore_index=True)


def load_gojet():
    t = read_table(DEBRIEFS / "GoJet Debriefs.xlsx", "Input")
    return pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Location"].astype(str).str.strip(),
        "Tail": t["Tail Number"],
        "IHC": t["Interior Heavy Clean (IHC)"].map(not_no),
        "RON": t["RON Clean (RON)"].map(not_no),
        "CE": t["Carpet Extraction (CE)"].map(not_no),
        "ED 1&2": [
            0 if a == "No" and b == "No" else 1
            for a, b in zip(t["Exterior Detail #1 (ED1)"], t["Exterior Detail #2 (ED2)"])
        ],
        "ED 3&4": 0,
    })


def load_mesa():
    t = read_table(DEBRIEFS / "Mesa Debriefs.xlsx", "Input")
    out = pd.DataFrame({"date": t["Date"].map(to_date),
                        "Tail": t["Tail Number"]})
    for flag, col in [
        ("IHC", "Interior Heavy Clean (IHC)"), ("RON", "RON Clean (RON)"),
        ("EC", "Exterior Clean (EC)"), ("ED", "Exterior Detail (ED)"),
        ("DSC", "Deep Seat Clean (DSC)"), ("CE", "Carpet Extraction (CE)"),
        ("ESS", "Disinfection (ESS)"), ("FDC", "Detailed Flight Deck Clean (Flight Deck)"),
    ]:
        out[flag] = t[col].map(yes) if col in t.columns else 0
    return out


def load_psa():
    t = read_table(DEBRIEFS / "PSA Debriefs.xlsx", "Input")
    svc7 = ["Interior Clean (I)", "Exterior Clean (E)", "Cockpit Cleaning (CC)",
            "Deep Seat Clean (DSC)", "Carpet Extraction (CE)",
            "Exterior Detail (ED1)", "Exterior Detail (ED2)"]
    combined = t[svc7].fillna("").astype(str).agg("".join, axis=1)
    return pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Location"].astype(str).str.strip().str[:3],
        "Tail": t["Tail Number"],
        "RON": (combined != "No" * 7).astype(int),
        "IC": t["Interior Clean (I)"].map(yes),
        "EC": t["Exterior Clean (E)"].map(yes),
        "CC": t["Cockpit Cleaning (CC)"].map(yes),
        "DSC": t["Deep Seat Clean (DSC)"].map(yes),
        "CE": t["Carpet Extraction (CE)"].map(yes),
        "ED1": t["Exterior Detail (ED1)"].map(yes),
        "ED2": t["Exterior Detail (ED2)"].map(yes),
        "ED3": t["Exterior Detail (ED3)"].isin(["Yes. 900", "Yes. 700"]).astype(int),
        "ED4": t["Exterior Detail (ED4)"].isin(["Yes. 700", "Yes. 900"]).astype(int),
        "Lav": t["Lav Tank Pressure Washing"].map(yes),
    })


def load_breeze():
    t = read_table(DEBRIEFS / "Breeze Debriefs.xlsx", "Table1")
    return pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Location"].astype(str).str.strip(),
        "Tail": t["Tail"],
        "Breeze RON": pd.to_numeric(t["Breeze RON"], errors="coerce").fillna(0).astype(int),
        "Breeze Ultra": pd.to_numeric(t["Breeze Ultra"], errors="coerce").fillna(0).astype(int),
    })


def load_ultra():
    t = read_table(DEBRIEFS / "Ultra Debriefs.xlsx", "Table1")
    out = pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Location"].astype(str).str.strip(),
        "Tail": t["Tail"],
        # One table holds two job types — "Ultra Cleaning" and "Shroud
        # Cleaning". Station specs filter on this so shroud jobs aren't
        # counted as Ultras (the workbook's COUNTIFS filtered Location only).
        "Service": t["Service"].astype(str).str.strip(),
    })
    shroud = next((c for c in t.columns if c and "Shroud" in str(c)), None)
    out["Shroud Count"] = (
        pd.to_numeric(t[shroud], errors="coerce").fillna(0) if shroud else 0
    )
    return out


def load_widebody():
    """Widebody turns (Sam, 2026-09-28) — one row per turn. Sheet1:
    Date/Name/Tail Number/Location/Customer/Widebody/Sub ID/Job Revenue.
    Plain sheet (no Excel table), so read_excel like the APU workbook."""
    t = pd.read_excel(DEBRIEFS / "Widebody Debriefs.xlsx", sheet_name="Sheet1")
    t.columns = [str(c).strip() for c in t.columns]
    out = pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Location"].astype(str).str.strip(),
        "Tail": t["Tail Number"],
    })
    return out[out["date"].notna()]


AA_JOB_TYPES = {"DTC": "AA Turn", "RSTC": "AA Turn", "RON": "AA RON",
                "RRON": "AA RON", "Security": "AA Security",
                "Ultra": "Ultra", "Shroud": "Shroud Cleaning"}


def load_aa():
    t = pd.read_csv(DEBRIEFS / "AA Debriefs.csv", dtype=str)
    t = t[t["Status"].str.strip() == "Completed"]
    svc = t["Job Type"].str.strip().map(AA_JOB_TYPES)
    out = pd.DataFrame({
        "date": pd.to_datetime(t["Job Date"], errors="coerce").dt.date,
        "Station": t["Station"].astype(str).str.strip(),
        "Tail": "",
        "Service": svc,
    })
    return out[out["Service"].notna()]


# APU Wash Sheet2 replaced its Status column with one numeric column per
# service (Aug 2026): 1 = completed, 0.5 = cancelled but half billable,
# 0 = no job. Only a 1 counts toward the pulse.
APU_SERVICES = ["APU Wash", "Landing Gear Bay Wash", "Partial Belly Wash"]


def load_apu():
    t = pd.read_excel(DEBRIEFS / "APU Wash.xlsx", sheet_name="Sheet2")
    t.columns = [str(c).strip() for c in t.columns]
    out = pd.DataFrame({
        "date": pd.to_datetime(t["Date"], errors="coerce").dt.date,
        "Location": t["Location"].astype(str).str.strip(),
    })
    for col in APU_SERVICES:
        if col not in t.columns:
            raise SystemExit(f"APU Wash.xlsx Sheet2 is missing column {col!r} "
                             f"(found {list(t.columns)})")
        out[col] = (pd.to_numeric(t[col], errors="coerce") == 1).astype(int)
    return out


# NetJets (FLL, Sam 2026-10-07): one row per job, a 0/1 column per service.
# Each service is priced per AIRFRAME on the workbook's own Pricing Sheet, so
# there is no single hours-per-job rate the way other programs have — every
# service earns its own price divided by a dollars-per-budgeted-hour rate.
# The loader prices each flagged service here (a "<col> $" column beside the
# count) so the per-service budget is the real money, not an average.
#
# Verified 2026-10-07: Job Revenue == the sum of those prices plus Other
# Price on all 102 rows, so this reproduces the workbook exactly rather than
# re-deriving it differently.
NETJETS_SERVICES = ["Standard Ex", "Complete Ex", "Standard Int",
                    "Complete Int", "Brightwork", "Extraction", "Aglaze",
                    "Deodorization", "Spot", "Other"]
# Debrief column -> the row it is priced on in the Pricing Sheet. "Other" is
# priced per job in the debrief's own "Other Price" column instead.
NETJETS_PRICED = {"Standard Ex": "Standard Exterior",
                  "Complete Ex": "Complete Exterior",
                  "Standard Int": "Standard Interior",
                  "Complete Int": "Complete Interior",
                  "Brightwork": "Brightwork", "Extraction": "Extraction",
                  "Aglaze": "Aglaze", "Deodorization": "Deodorization",
                  "Spot": "Spot"}


def _netjets_prices(path):
    """{service: {aircraft type: price}} from the workbook's Pricing Sheet."""
    pr = pd.read_excel(path, sheet_name="Pricing Sheet")
    pr.columns = [str(c).strip() for c in pr.columns]
    key = pr.columns[0]
    types = [c for c in pr.columns[1:] if c and not c.startswith("Unnamed")]
    out = {}
    for _, r in pr.iterrows():
        svc = str(r[key]).strip()
        if svc and svc.lower() != "nan":
            out[svc] = {t: pd.to_numeric(r[t], errors="coerce") for t in types}
    return out


def load_netjets():
    path = DEBRIEFS / "NetJets Debriefs.xlsx"
    prices = _netjets_prices(path)
    t = pd.read_excel(path, sheet_name="Debriefs")
    t.columns = [str(c).strip() for c in t.columns]
    ac = t["AC Type"].astype(str).str.strip()
    out = pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": "FLL",          # the workbook is FLL-only by construction
        "Tail": t["Tail"],
        "Job Revenue": pd.to_numeric(t["Job Revenue"], errors="coerce").fillna(0),
    })
    unpriced = set()
    for col in NETJETS_SERVICES:
        if col not in t.columns:
            raise SystemExit(f"NetJets Debriefs is missing column {col!r} "
                             f"(found {list(t.columns)})")
        done = (pd.to_numeric(t[col], errors="coerce").fillna(0) > 0)
        out[col] = done.astype(int)
        if col == "Other":
            money = pd.to_numeric(t.get("Other Price"), errors="coerce").fillna(0)
        else:
            tbl = prices.get(NETJETS_PRICED[col], {})
            money = ac.map(lambda a: tbl.get(a))
            unpriced |= {(col, a) for a, m in zip(ac[done], money[done])
                         if pd.isna(m)}
            money = pd.to_numeric(money, errors="coerce").fillna(0)
        out[col + " $"] = money.where(done, 0.0)
    if unpriced:
        print(f"  !! NetJets: no Pricing Sheet entry for {sorted(unpriced)} — "
              f"those jobs earn 0 budget hours")
    out = out[out["date"].notna()]
    gap = float((out[[c + " $" for c in NETJETS_SERVICES]].sum(axis=1)
                 - out["Job Revenue"]).abs().sum())
    print(f"  NetJets: priced {len(out)} job(s); per-service total vs the "
          f"workbook's Job Revenue differs by ${gap:,.2f}")
    return out


# ── Private/MRO locations as pulse stations (Sam, 2026-10-11) ───────────────
#
# An MRO location is not a different kind of thing: it has budgeted and worked
# hours per day exactly as a commercial station does. Only the SHAPE of its
# config differs — the budget comes from a drag-and-drop forecast rather than
# debrief counts.
#
# Rather than teach the build a second kind of station, each scheduled job is
# exploded into one row per working day in a synthetic `MRO_Jobs` table:
# date / Location / Service / Hours. That makes an MRO service an ordinary
# counted service — COUNTIFS gives jobs in progress that day, and `hours_from`
# (built for NetJets) sums the priced hours beside it. Everything downstream —
# weeks, KPIs, variance, the overview — then works untouched.
#
# PARITY NOTE: the job -> hours rule is implemented three times now (here,
# engine/mro.py's job_hours, and mro.js's client-side mroRate). They must
# agree: price / the service's revenue rate, premium services on the premium
# goal, spread evenly across the job's working days. A test below compares
# this build's totals against the platform's for the same window.
MRO_WORKBOOK = "Private and MRO Pulse Locations.xlsx"
MRO_SCHEDULES = "MRO Schedules"


def _mro_configs():
    """One sheet per location; A/B is a key/value block read until the first
    blank key (same contract engine/mro.py reads)."""
    wb = openpyxl.load_workbook(BUDGETS_DIR / MRO_WORKBOOK, read_only=True,
                                data_only=True)
    out = {}
    for ws in wb.worksheets:
        kv = {}
        for r in range(1, ws.max_row + 1):
            k = ws.cell(row=r, column=1).value
            if k is None or str(k).strip() == "":
                break
            kv[re.sub(r"\s+", " ", str(k)).strip().lower()] = ws.cell(row=r, column=2).value
        dists = [d.strip() for d in str(kv.get("labor distribution") or "").split(",") if d.strip()]
        if not dists:
            continue
        def num(key):
            try:
                return float(kv.get(key) or 0)
            except (TypeError, ValueError):
                return 0.0
        out[ws.title.strip()] = {
            "labor_dists": dists, "hourly_goal": num("hourly goal"),
            "facility_hours": num("facility hours"),
            "premium_goal": num("premium hourly goal"),
            "premium_services": [x.strip() for x in
                                 str(kv.get("premium services") or "").split(",") if x.strip()],
            "attribution": str(kv.get("attribution") or "").strip().lower(),
        }
    wb.close()
    return out


def _mro_working_days(job):
    """ISO dates the job actually works: [start, end] minus `skip`."""
    try:
        a = date.fromisoformat(str(job.get("start") or ""))
        b = date.fromisoformat(str(job.get("end") or ""))
    except ValueError:
        return []
    if b < a:
        return []
    skip = set(job.get("skip") or [])
    out, d = [], a
    while d <= b:
        if d.isoformat() not in skip:
            out.append(d)
        d += timedelta(days=1)
    return out


def _mro_job_rate(job, cfg):
    """The revenue rate a job's hours divide by — premium services carry the
    premium goal; custom jobs and everything else the base."""
    pg = cfg.get("premium_goal") or 0
    if pg and not job.get("custom") and job.get("service") in (cfg.get("premium_services") or []):
        return pg
    return cfg.get("hourly_goal") or 0


def load_mro_jobs(cfgs):
    """Every scheduled job exploded to one row per working day."""
    rows = []
    for loc, cfg in cfgs.items():
        path = BUDGETS_DIR / MRO_SCHEDULES / f"{loc}.json"
        if not path.exists():
            print(f"  !! MRO {loc}: no schedule file — budget will be the "
                  f"facility allowance only")
            continue
        doc = json.loads(path.read_text(encoding="utf-8"))
        for job in doc.get("jobs", []):
            wd = _mro_working_days(job)
            rate = _mro_job_rate(job, cfg)
            try:
                price = float(job.get("price"))
            except (TypeError, ValueError):
                continue
            if not wd or not rate:
                continue
            # Store the MONEY and the rate it is earned at, not pre-divided
            # hours: the service row then carries the real revenue rate, so
            # the rates card can say "$44 of revenue / hour" instead of a
            # meaningless placeholder, and hours = revenue / rate falls out
            # of the same `hours_from` machinery NetJets uses.
            per = price / len(wd)
            svc = (job.get("service") or job.get("name") or "Custom job").strip()
            is_custom = bool(job.get("custom")) or not job.get("service")
            for d in wd:
                rows.append({"date": d, "Location": loc, "Service": svc,
                             "Rate": rate, "Revenue": round(per, 4),
                             "Custom": int(is_custom)})
    t = pd.DataFrame(rows, columns=["date", "Location", "Service", "Rate",
                                    "Revenue", "Custom"])
    print(f"  MRO_Jobs: {len(t)} job-day row(s) across {t['Location'].nunique() if len(t) else 0} location(s)")
    return t


def _mro_group_map(loc):
    """{SERVICE NAME (upper): group} plus the group order, for one location.
    Per-location entries win over the universal ones."""
    raw = json.loads((HERE / "mro_groups.json").read_text(encoding="utf-8"))
    order, by_svc, custom_group = [], {}, None
    for src in (raw.get("_universal") or {}, raw.get(loc) or {}):
        for g, spec in src.items():
            if g not in order:
                order.append(g)
            if spec.get("custom"):
                custom_group = g
            for n in spec.get("services") or []:
                by_svc[n.strip().upper()] = g
    # per-location wins: re-apply its mapping last
    for g, spec in (raw.get(loc) or {}).items():
        for n in spec.get("services") or []:
            by_svc[n.strip().upper()] = g
    return by_svc, custom_group, order


def mro_stations(cfgs, jobs):
    """Each MRO location as an ordinary pulse station."""
    out = {}
    for loc, cfg in cfgs.items():
        svcs = []
        # Service names must be grouped CASE-INSENSITIVELY, because spec_mask
        # compares upper-cased: two rows for "PAXX DOOR" and "PAXX Door" would
        # each match the other's jobs and double-count them. Hand-typed job
        # names collide like this routinely — TUS MHI carries three spellings
        # of one door-cleaning job. The most common spelling wins the label.
        seen, rates, custom_jobs = {}, {}, {}
        if len(jobs):
            mine = jobs[jobs["Location"] == loc]
            for n, r, cu in zip(mine["Service"], mine["Rate"], mine["Custom"]):
                seen.setdefault(n.upper(), []).append(n)
                rates.setdefault(n.upper(), set()).add(float(r))
                custom_jobs[n.upper()] = custom_jobs.get(n.upper(), 0) or int(cu)
        gmap, custom_group, gorder = _mro_group_map(loc)
        used_groups = []
        for key in sorted(seen):
            label = max(set(seen[key]), key=seen[key].count)
            group = gmap.get(key)
            if group is None and custom_jobs.get(key):
                group = custom_group     # every custom job lands in one row
            if group and group not in used_groups:
                used_groups.append(group)
            # One name can legitimately carry two rates — PVU's XZILON
            # APPLICATION is premium on booked jobs and base on custom ones.
            # Split those into a row each, keyed on the rate so neither can
            # match the other's jobs, and say which is which.
            for rate in sorted(rates[key], reverse=True):
                crit = [["Location", loc], ["Service", label], ["Rate", rate]]
                name = label if len(rates[key]) == 1 else (
                    f"{label} (premium)" if rate == (cfg.get("premium_goal") or 0)
                    else f"{label} (standard rate)")
                svcs.append({
                    "name": name, "kind": "count", "aircraft": False,
                    "rate": rate, "group": group,
                    "specs": [{"fn": "COUNTIFS", "table": "MRO_Jobs",
                               "sum_col": None, "criteria": crit}],
                    "hours_from": [{"fn": "SUMIFS", "table": "MRO_Jobs",
                                    "sum_col": "Revenue", "criteria": crit}],
                })
        if cfg["facility_hours"]:
            svcs.append({"name": "Facility Budget", "kind": "fixed",
                         "aircraft": False, "rate": cfg["facility_hours"]})
        # MRO crews work day shifts: plain calendar day unless the sheet says
        # otherwise (the same `Attribution` knob build_mro_hours.py honours).
        plain = [] if cfg["attribution"] == "shift" else list(cfg["labor_dists"])
        out[loc] = {"labor_keys": cfg["labor_dists"],
                    "salary_keys": cfg["labor_dists"],
                    "fac_only": not seen, "facility": False,
                    "hours_from_first_debrief": False,
                    "labor_from_first_debrief": None, "shift_window": None,
                    "plain_day_dists": plain,
                    "group_order": [g for g in gorder if g in used_groups],
                    "vertical": "mro", "services": svcs}
    return out


JSX_SERVICES = ["RON", "Interior Detail", "Exterior Detail", "Carpet Extraction"]
# Foxtrot took over the JSX contract on 2026-08-01. Rows dated before that are
# backfilled service history from the previous vendor (they all carry no
# revenue) — they are not work we performed, so they never count.
JSX_START = date(2026, 8, 1)


def load_jsx():
    """JSX Debriefs 'Debriefs' table: one row per tail serviced, with a 0/1
    column per service and the airport in 'Service Location'."""
    t = read_table(DEBRIEFS / "JSX Debriefs.xlsx", "Debriefs")
    out = pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Service Location"].astype(str).str.strip().str.upper(),
        "Tail": t["Tail Number"],
        "Plane Type": t["Plane Type"].astype(str).str.strip(),
    })
    for col in JSX_SERVICES:
        if col not in t.columns:
            raise SystemExit(f"JSX Debriefs is missing column {col!r} "
                             f"(found {list(t.columns)})")
        out[col] = (pd.to_numeric(t[col], errors="coerce").fillna(0) == 1).astype(int)

    # Exterior details are priced by airframe (Aug 2026): ATR or, for anything
    # else, the 145 rate. "Anything else" is deliberate per the owner — an
    # unreadable Plane Type (the debrief carries the odd "#N/A") lands in 145
    # rather than being dropped.
    is_atr = out["Plane Type"].str.upper().str.contains("ATR", na=False)
    ed = out["Exterior Detail"] == 1
    out["Exterior Detail ATR"] = (ed & is_atr).astype(int)
    out["Exterior Detail 145"] = (ed & ~is_atr).astype(int)
    odd = sorted(set(out.loc[ed & ~is_atr, "Plane Type"]) - {"EMB 145", "EMB 135"})
    if odd:
        print(f"  JSX: exterior details with unrecognised Plane Type counted "
              f"as 145: {odd}")
    before = len(out)
    out = out[[d is not None and d >= JSX_START for d in out["date"]]]
    print(f"  JSX: dropped {before - len(out)} pre-{JSX_START} row(s) "
          f"(pre-contract service history)")
    return out


def load_frontier():
    t = read_table(DEBRIEFS / "Frontier Debriefs.xlsx", "Table1")
    t = t[t["Status"] == "Complete"]
    return pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Location"].astype(str).str.strip(),
        "Tail": t["Tail Number"],
        "Aircraft Type and Service": t["Aircraft Type and Service"],
    })


# ------------------------------------------------------- closeout compare

# Closeout Compare reconciles each night's closeout against the debriefs. The
# pulse acts on exactly two discrepancy types (Sam, 2026-09-08); everything
# else in that sheet (service mismatches, typos, missing-from-closeout) is a
# paperwork matter that does not change a debrief count.
CLOSEOUT_DOUBLE = "double debrief"          # counted twice -> corrected here
CLOSEOUT_MISSING = "missing from debrief"   # never submitted -> NOT invented
# Program -> the debrief table the pulse counts it from, so a discrepancy is
# only pinned on stations that actually read that program.
CLOSEOUT_PROGRAM_TABLE = {
    "envoy": "Envoy_Debriefs", "dfw": "Envoy_Debriefs",
    "regional": "Envoy_Debriefs", "psa": "PSA_Debriefs",
    "gojet": "GoJet_Debriefs", "mesa": "Mesa_Debriefs",
    "ultra": "Ultra_Debriefs", "breeze": "Breeze_Debriefs",
    "widebody": "Widebody_Debriefs",
    "jsx": "JSX_Debriefs", "frontier": "Frontier_Debriefs",
}


def _closeout_kind(text):
    """Canonical label only. The sheet also carries free-text descriptions of
    the same situations ("Double debrief flagged", "Missing in debrief (on
    closeout...)"); those are left alone rather than parsed, because a wrong
    match here would silently change a station's counts."""
    t = re.sub(r"\s+", " ", str(text or "")).strip().lower()
    if t == CLOSEOUT_DOUBLE:
        return "double"
    if t == CLOSEOUT_MISSING:
        return "missing"
    return None


# Sheet holding the closeout-vs-debrief discrepancies. Renamed from "Sheet1"
# on 2026-09-16, which failed every refresh for the day until it was caught;
# the old name is kept as a fallback and an unknown name now fails with the
# workbook's actual sheet list instead of a bare pandas ValueError. The
# workbook's other tab, "Work Order Findings", is a different schema (Finding,
# not Discrepancy) and is deliberately not read here.
CLOSEOUT_SHEETS = ("Closeout Findings", "Sheet1")


def _closeout_sheet(path):
    have = pd.ExcelFile(path).sheet_names
    for name in CLOSEOUT_SHEETS:
        if name in have:
            return name
    raise SystemExit(
        f"Closeout Compare.xlsx has none of {list(CLOSEOUT_SHEETS)} — its "
        f"sheets are {have}. Add the new name to CLOSEOUT_SHEETS if it was "
        f"renamed again.")


def load_closeout():
    """Open Missing-from-Debrief / Double-Debrief rows, by (station, date).

    A row whose Status is Closed (any case) is settled: the debrief is then
    authoritative, so it neither corrects a count nor raises an asterisk."""
    path = DEBRIEFS / "Closeout Compare.xlsx"
    t = pd.read_excel(path, sheet_name=_closeout_sheet(path))
    t.columns = [str(c).strip() for c in t.columns]
    rows, closed_n, other = [], 0, 0
    for _, r in t.iterrows():
        kind = _closeout_kind(r.get("Discrepancy"))
        if kind is None:
            other += 1
            continue
        if str(r.get("Status") or "").strip().lower() == "closed":
            closed_n += 1
            continue
        d = to_date(r.get("Date"))
        if d is None:
            continue
        rows.append({
            "date": d,
            "loc": str(r.get("Location") or "").strip().upper()[:3],
            "tail": str(r.get("Tail") or "").strip().upper(),
            "program": str(r.get("Program") or "").strip().lower(),
            "table": CLOSEOUT_PROGRAM_TABLE.get(
                str(r.get("Program") or "").strip().lower()),
            "kind": kind,
        })
    out = pd.DataFrame(rows, columns=["date", "loc", "tail", "program",
                                      "table", "kind"])
    n_d = int((out["kind"] == "double").sum()) if len(out) else 0
    n_m = int((out["kind"] == "missing").sum()) if len(out) else 0
    print(f"  Closeout Compare: {n_d} open double, {n_m} open missing "
          f"({closed_n} closed, {other} other discrepancy types ignored)")
    return out


def apply_double_debriefs(tables, closeout):
    """Drop the duplicate row an open Double Debrief refers to.

    Only ever removes one row per flagged job, and only when the table really
    does hold two or more matching rows — if someone has already deleted the
    duplicate, correcting again would under-count."""
    if closeout.empty:
        return 0
    fixed = 0
    for _, r in closeout[closeout["kind"] == "double"].iterrows():
        tname = r["table"]
        t = tables.get(tname)
        if t is None or not r["tail"]:
            continue
        m = (t["date"] == r["date"]) &             (t["Tail"].astype(str).str.strip().str.upper() == r["tail"])
        loc_col = "Station" if "Station" in t.columns else (
            "Location" if "Location" in t.columns else None)
        if loc_col and r["loc"]:
            m &= t[loc_col].astype(str).str.strip().str.upper().str[:3] == r["loc"]
        idx = list(t.index[m])
        if len(idx) >= 2:
            tables[tname] = t.drop(index=idx[-1])
            fixed += 1
    print(f"  double debriefs corrected (one duplicate row dropped each): {fixed}")
    return fixed


# ---------------------------------------------------------------- paylocity

NORM = lambda s: re.sub(r"[^a-z0-9]", "", str(s).lower())


def load_budget_workbook():
    """Service Budgets.xlsx: one sheet per location, Service | Budget rows.
    This workbook is the authority on which locations exist, which services
    each has, and the hours-per-job rates."""
    wb = openpyxl.load_workbook(BUDGETS_DIR / "Service Budgets.xlsx", data_only=True)
    out = {}
    for ws in wb.worksheets:
        rows = []
        for r in range(2, ws.max_row + 1):
            svc, rate = ws.cell(row=r, column=1).value, ws.cell(row=r, column=2).value
            if svc is None or str(svc).strip() == "":
                continue
            rows.append((str(svc).strip(), float(rate) if rate is not None else 0.0))
        if rows:
            out[ws.title.strip()] = rows
    wb.close()
    return out


def merge_budget_config(stations, budgets, catalog, overrides):
    """Build the effective per-station config: the budget workbook decides the
    station list, service list, and rates; specs come from station_overrides
    first, then stations.json for pre-existing station+service pairs, then the
    service catalog templates ({LOC} = first 3 letters of the sheet name).
    Unknown service names fall back to a flat daily budget, with a warning.
    An override may carry "display" to show the sheet under another name."""
    merged = {}
    for st_name, rows in budgets.items():
        code = st_name.strip()[:3].upper()
        base = stations.get(st_name)
        ovr = overrides.get(st_name, {})
        ovr_services = ovr.get("services", {})
        base_by_name = ({NORM(s["name"]): (i, s) for i, s in enumerate(base["services"])}
                        if base else {})
        base_aircraft = ({r - 3 for r in base["aircraft_service_rows"]} if base else set())
        services = []
        for svc_name, rate in rows:
            key = NORM(svc_name)
            if key in ovr_services:
                svc = dict(ovr_services[key])
            elif key in base_by_name:
                i, src = base_by_name[key]
                svc = {k: v for k, v in src.items() if k != "sheet_row"}
                svc["aircraft"] = (not base["fac_only"]) and i in base_aircraft
            elif key in catalog:
                tpl = catalog[key]
                svc = {"kind": tpl["kind"], "aircraft": tpl["aircraft"]}
                if tpl["kind"] == "count":
                    svc["specs"] = [
                        {"fn": sp["fn"], "table": sp["table"], "sum_col": sp["sum_col"],
                         "criteria": [[c, (code if v == "{LOC}" else v)]
                                      for c, v in sp["criteria"]]}
                        for sp in tpl["specs"]]
            else:
                print(f"  !! {st_name}: unknown service {svc_name!r} — "
                      f"treating as flat daily budget")
                svc = {"kind": "fixed", "aircraft": False}
            svc["name"], svc["rate"] = svc_name, rate
            services.append(svc)
        fac_only = all(s["kind"] == "fixed" for s in services)
        if "labor_keys" in ovr:
            labor_keys = ovr["labor_keys"]
            salary_keys = ovr.get("salary_keys", labor_keys)
        elif base:
            labor_keys, salary_keys = base["labor_keys"], base["salary_keys"]
        else:
            k = f"{code} FAC" if "FAC" in st_name.upper() else f"{code} CABIN"
            labor_keys = salary_keys = [k]
            print(f"  new location {st_name!r}: labor dist {k!r} (convention)")
        # Facility stations take plain day attribution (no overnight shift-back).
        # Name test first so a facility that gains a counted service keeps the
        # right treatment; fac_only covers any future non-"FAC" naming.
        facility = st_name.strip().upper().endswith("FAC") or fac_only
        # The tab takes the sheet's name unless the override renames it
        # (Sam, 2026-09-28: the "DFW WIDEBODY" sheet shows as "DFW WIDE").
        # Renamed HERE, so everything downstream — debrief gates, manager
        # grouping, the month payload — speaks the one name. The 3-letter
        # airport code must survive the rename; both forms start "DFW".
        merged[ovr.get("display", st_name)] = {
                           "labor_keys": labor_keys, "salary_keys": salary_keys,
                           "fac_only": fac_only, "facility": facility,
                           "hours_from_first_debrief":
                               bool(ovr.get("hours_from_first_debrief")),
                           "labor_from_first_debrief":
                               ovr.get("labor_from_first_debrief"),
                           "shift_window": ovr.get("shift_window"),
                           "plain_day_dists": ovr.get("plain_day_dists"),
                           # Collapsible service groups: the order they render
                           # in. A station without this renders flat.
                           "group_order": ovr.get("group_order"),
                           "services": services}
    for st in stations:
        if st not in budgets:
            print(f"  !! {st} is in stations.json but has no Service Budgets "
                  f"sheet — dropped from the pulse")
    return merged


def load_managers(station_names):
    """Airport code (first 3 letters, both sides) → managers, per Location
    Management.csv. A station lands in every matching manager's group."""
    m = pd.read_csv(LOCMGMT_DIR / "Location Management.csv", dtype=str,
                    encoding="utf-8-sig")
    by_code = {}
    for loc, mgr in zip(m["Location"], m["Manager"]):
        if pd.isna(loc) or pd.isna(mgr):
            continue
        by_code.setdefault(loc.strip()[:3].upper(), set()).add(mgr.strip())
    return {st: sorted(by_code.get(st.strip()[:3].upper(), set()))
            for st in station_names}


def _norm_eid(s):
    s = str(s or "").strip()
    if s[:1] in ("A", "a"):
        s = s[1:]
    return s.lstrip("0") or ("0" if s else "")


def load_early_terms():
    """Employee id -> Termination Form date. The form beats Paylocity, which
    keeps leavers 'Active' until the pay cycle closes — without this cap a
    salaried leaver keeps imputing hours for days after they're gone."""
    path = DEFINITIVE_DIR / "Early Terminations.csv"
    out = {}
    if not path.exists():
        return out
    for _, r in pd.read_csv(path, dtype=str).iterrows():
        try:
            out[_norm_eid(r["Employee Id"])] = datetime.strptime(
                (r["Form Date"] or "").strip(), "%Y-%m-%d").date()
        except (KeyError, TypeError, ValueError):
            continue
    return out


def load_pay_type_changes():
    """Employee id -> date their new Salary pay type actually starts.

    Paylocity flips Pay Type Code the day a promotion is entered, but the
    person keeps being paid hourly until the new pay period. Without this the
    pulse back-applies the promotion across the whole month: it drops every
    punch they actually worked (only Hourly punches count) AND adds 40/7 of
    imputed salaried time a day (Sam, 2026-09-23, Riley Pinterich / LIT).

    Maintained by the roster job (core: pay_type_changes.maintain), which
    detects the flip and dates it to the start of the next pay period; rows
    can also be hand-added or their Effective date corrected to whatever
    payroll actually agreed. Only the hourly->salary direction is deferred.
    """
    path = DEFINITIVE_DIR / "Pay Type Changes.csv"
    out = {}
    if not path.exists():
        return out
    for _, r in pd.read_csv(path, dtype=str).iterrows():
        if str(r.get("To") or "").strip().lower() != "salary":
            continue
        try:
            out[_norm_eid(r["Employee Id"])] = datetime.strptime(
                (r["Effective"] or "").strip(), "%Y-%m-%d").date()
        except (KeyError, TypeError, ValueError):
            continue
    return out


def load_employees():
    e = pd.read_csv(PAYLOCITY / "Basic Employee Info.csv", dtype=str, encoding="cp1252")
    e = e[~e["Job Title"].isin(EXCLUDED_TITLES)].copy()
    term = pd.to_datetime(e["Termination Date"], errors="coerce")
    active = e["Employee Status Description"].str.strip() == "Active"
    early = load_early_terms()
    last_days, capped = [], 0
    for a, t, i in zip(active, term, e["Employee Id"]):
        if a:
            form = early.get(_norm_eid(i))
            if form is not None:
                last_days.append(min(TODAY, form))
                capped += 1
            else:
                last_days.append(TODAY)
        else:
            last_days.append(t.date() if pd.notna(t) else None)
    e["last_day"] = last_days
    if capped:
        print(f"early terminations: capped last_day for {capped} "
              f"form-termed people Paylocity still shows Active")
    e["labor_dist"] = e["Labor Dist Description"].fillna("").str.strip()
    e["pay_type"] = e["Pay Type Code"].fillna("").str.strip()
    e["id"] = e["Employee Id"].str.strip()
    # The mirror image of last_day: last_day caps when someone stops counting
    # as salaried, salary_from sets when they start. Before it they are still
    # hourly in fact, so their punches count and they impute nothing.
    pend = load_pay_type_changes()
    e["salary_from"] = [
        pend.get(_norm_eid(i)) if p == "Salary" else None
        for i, p in zip(e["Employee Id"], e["pay_type"])
    ]
    deferred = int(sum(1 for v in e["salary_from"] if v is not None))
    if deferred:
        who = ", ".join(
            f"{f} {l} -> Salary {v:%Y-%m-%d}"
            for f, l, v in zip(e["First Name"], e["Last Name"], e["salary_from"])
            if v is not None)
        print(f"pay type changes: {deferred} promotion(s) not yet in effect, "
              f"still counted hourly ({who})")
    return e


def hours_file_updated():
    """When This Years Hours.csv was last regenerated (naive local/Eastern).
    CI: SharePoint lastModifiedDateTime captured by fetch_sources.py;
    local: file mtime."""
    meta = PAYLOCITY / "sources_meta.json"
    if meta.exists():
        import json as _json
        ts = _json.loads(meta.read_text()).get("This Years Hours.csv")
        if ts:
            from zoneinfo import ZoneInfo
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
            return dt.astimezone(ZoneInfo("America/New_York")).replace(tzinfo=None)
    return datetime.fromtimestamp((PAYLOCITY / "This Years Hours.csv").stat().st_mtime)


def load_hours(emp):
    h = pd.read_csv(PAYLOCITY / "This Years Hours.csv", dtype=str, encoding="cp1252")
    # Paylocity renamed this column in the Aug 2026 export; accept either name
    if "Labor Dist Name" not in h.columns and "Cost Center 2 Name" in h.columns:
        h = h.rename(columns={"Cost Center 2 Name": "Labor Dist Name"})
    h["work_date"] = pd.to_datetime(h["Work Date"], errors="coerce").dt.date
    h = h[(h["work_date"] >= WINDOW_START) & (h["work_date"] <= TODAY)].copy()
    h["hours"] = pd.to_numeric(h["Summation of Paid Duration (hours)"], errors="coerce").fillna(0)
    punch = pd.to_datetime(h["Punch In Time"], errors="coerce")
    punched_out = h["Punch Out Time"].fillna("").str.strip() != ""

    # Holiday/PTO credits: punch-in placeholder, no punch-out, but paid hours
    # already posted. Excluded from worked hours (owner decision, Aug 2026).
    holiday = punch.notna() & ~punched_out & (h["hours"] > 0)
    h = h[~holiday].copy()
    punch = punch[~holiday]
    punched_out = punched_out[~holiday]
    print(f"  holiday/PTO rows excluded: {int(holiday.sum())}")

    # Two day attributions, chosen per station in build_month():
    #  attr_date       — workbook helper DAY/MONTH(punch-in − 12h). Commercial
    #                    stations work overnight, so a shift credits the day it
    #                    started.
    #  attr_date_plain — the calendar day the punch actually falls in. Facility
    #                    stations don't run overnight shifts, so shifting back
    #                    12 h would push a normal day shift onto the day before
    #                    (owner decision, Aug 2026).
    # Rows with no punch-in fall back to Work Date in both.
    shifted = punch - timedelta(hours=12)
    h["attr_date"] = [
        s.date() if pd.notna(s) else w for s, w in zip(shifted, h["work_date"])
    ]
    h["attr_date_plain"] = [
        p.date() if pd.notna(p) else w for p, w in zip(punch, h["work_date"])
    ]
    # Local clock hour the shift began, for stations split into day/night
    # crews sharing one labor dist (CLE DAY / CLE NIGHT). Punch times are
    # already station-local in the Paylocity export. -1 = no punch-in.
    h["punch_hour"] = [int(p.hour) if pd.notna(p) else -1 for p in punch]

    # Shifts still in progress when the CSV was generated (punch-in, no
    # punch-out, 0 paid hours). The pulse no longer estimates them (Sam,
    # 2026-09-08): they carry 0 paid hours so they contribute nothing, and
    # the day is marked with an asterisk saying how many are outstanding.
    # Only punches from the 18 h before the file was generated count as "in
    # progress" - older opens are abandoned punches rather than tonight's
    # crew, and future-dated placeholder rows are not shifts at all.
    updated = hours_file_updated()
    age_h = (updated - punch).dt.total_seconds() / 3600
    h["incomplete"] = (punch.notna() & ~punched_out & (h["hours"] == 0)
                       & (age_h >= 0) & (age_h <= 18)).astype(int)
    print(f"  incomplete punches (excluded, flagged on the page): "
          f"{int(h['incomplete'].sum())} "
          f"(hours file generated {updated:%b %d %I:%M %p})")
    id2dist = dict(zip(emp["id"], emp["labor_dist"]))
    id2pay = dict(zip(emp["id"], emp["pay_type"]))
    eid = h["Employee Id"].str.strip()
    raw = h["Labor Dist Name"].fillna("").str.strip()
    h["labor_dist"] = [
        id2dist.get(i, "") if r == "Not Defined" else r for i, r in zip(eid, raw)
    ]
    h["pay_type"] = eid.map(id2pay).fillna("")
    # A promotion that has not taken effect yet: these punches are hourly work
    # and must be counted as such (see load_pay_type_changes).
    id2from = {i: v for i, v in zip(emp["id"], emp["salary_from"]) if v is not None}
    if id2from:
        # map() yields NaN, not None, for everyone not on the list
        pre = [pd.notna(sf) and wd is not None and wd < sf
               for sf, wd in zip(eid.map(id2from), h["work_date"])]
        h.loc[pre, "pay_type"] = "Hourly"
        print(f"  punch rows re-counted as hourly (promotion not yet in "
              f"effect): {int(sum(pre))}")
    return h


# ---------------------------------------------------------------- compute

def spec_mask(df, spec):
    m = pd.Series(True, index=df.index)
    for col, val in spec["criteria"]:
        if col not in df.columns:
            return pd.Series(False, index=df.index)
        if isinstance(val, str):
            m &= df[col].astype(str).str.strip().str.upper() == val.strip().upper()
        else:
            m &= df[col] == val
    return m


def month_days(year, month):
    return monthrange(year, month)[1]


DOW_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def service_start(svc):
    """A service's "active as of" date, or None.

    A contract line can start mid-window (CVG's Envoy Facility begins
    2026-09-14). Before that date the service contributes nothing at all
    rather than being back-applied over history."""
    raw = svc.get("starts")
    if not raw:
        return None
    if isinstance(raw, date):
        return raw
    return datetime.strptime(str(raw).strip(), "%Y-%m-%d").date()


def dow_daily_rates(svc):
    """Per-weekday hours for a fixed service that only runs on set days.

    FLL FAC's facility budget and STL AA CAB's mail run are the same shape -
    set hours on set weekdays - so they share one mechanism:

      dow_days       weekdays the workbook rate applies to; every other day
                     is 0. (Mail Running: 8 h Mon-Fri, nothing at weekends.)
      dow_overrides  per-day hours replacing the rate on named days, for a
                     day that runs at a different level. (FLL FAC: the 2 h
                     Saturday half-shift.)

    The workbook rate keeps its plain meaning - "hours on a normal working
    day" - so editing Service Budgets.xlsx still moves the budget. Returns a
    list indexed by date.weekday(), or None for a flat every-day rate.
    """
    days_on = svc.get("dow_days")
    over = svc.get("dow_overrides") or {}
    if not days_on and not over:
        return None
    for key, names in (("dow_days", days_on or []), ("dow_overrides", over)):
        unknown = set(names) - set(DOW_NAMES)
        if unknown:
            raise SystemExit(f"{key} for {svc.get('name')!r} has unknown "
                             f"day(s) {sorted(unknown)}; use {DOW_NAMES}")
    rate = float(svc.get("rate") or 0)
    return [float(over[n]) if n in over
            else (rate if days_on and n in days_on else 0.0)
            for n in DOW_NAMES]


def aa_missing_days(aa_full, year, month, upto_day):
    """Elapsed days on which a station has no AA rows at all.

    AA's feed runs through AA's internal system and does not update on its own,
    so whole days go absent. The pulse used to fill those with a 7-day average;
    it no longer invents data (Sam, 2026-09-08) — the day simply counts what
    arrived, and the page marks it so nobody reads a gap as a slow night.
    Returns {station code: {day numbers}}."""
    out = {}
    if aa_full is None or aa_full.empty:
        return out
    for stn, grp in aa_full.groupby("Station"):
        have = {d.day for d in grp["date"]
                if d is not None and d.year == year and d.month == month}
        missing = {d for d in range(1, upto_day + 1) if d not in have}
        if missing:
            out[str(stn).strip().upper()[:3]] = missing
    return out


def first_debrief_dates(stations, tables):
    """Earliest date each flagged station has a counted service on.

    Stations carrying `hours_from_first_debrief` only start counting labor from
    their first real debrief: the hours before that are implementation and
    training for a contract that hadn't started, and shouldn't score against
    the station. Derived from the debrief data (not a hardcoded date) so it
    corrects itself as debriefs land or are backfilled.
    """
    out = {}
    for st_name, cfg in stations.items():
        if not cfg.get("hours_from_first_debrief"):
            continue
        best = None
        for svc in cfg["services"]:
            for spec in svc.get("specs", []):
                t = tables.get(spec["table"])
                if t is None or t.empty:
                    continue
                m = spec_mask(t, spec)
                if spec["fn"] == "SUMIFS":
                    m &= pd.to_numeric(t[spec["sum_col"]],
                                       errors="coerce").fillna(0) > 0
                dates = [d for d in t.loc[m, "date"] if d is not None]
                if dates and (best is None or min(dates) < best):
                    best = min(dates)
        out[st_name] = best
        if best:
            print(f"  {st_name}: labor counted from first debrief {best}")
        else:
            print(f"  !! {st_name}: no debriefs at all — all labor suppressed")
    return out


def labor_gate_dates(stations, tables):
    """Per-labor-dist start dates, for a station that pools several cost
    centres where only some had a launch.

    DFW-DAL is the case: DFW has traded for months, but DAL only went live
    with JSX on its first debrief, and the labour charged to DAL PRIV before
    that (mostly imputed salaried time) is implementation, not performance.
    Gating the whole station would erase DFW's real history, so the station
    names which labor dist is gated and which services date its launch:

        "labor_from_first_debrief": {"DAL PRIV": ["jsxron", ...]}

    Returns {station: {LABOR DIST: date}}; dates are derived, not hardcoded.
    """
    out = {}
    for st_name, cfg in stations.items():
        rules = cfg.get("labor_from_first_debrief") or {}
        if not rules:
            continue
        by_key = {NORM(s["name"]): s for s in cfg["services"]}
        gates = {}
        for dist, svc_keys in rules.items():
            best = None
            for k in svc_keys:
                svc = by_key.get(NORM(k))
                if not svc:
                    print(f"  !! {st_name}: labor gate names unknown service {k!r}")
                    continue
                for spec in svc.get("specs", []):
                    t = tables.get(spec["table"])
                    if t is None or t.empty:
                        continue
                    m = spec_mask(t, spec)
                    if spec["fn"] == "SUMIFS":
                        m &= pd.to_numeric(t[spec["sum_col"]],
                                           errors="coerce").fillna(0) > 0
                    dates = [d for d in t.loc[m, "date"] if d is not None]
                    if dates and (best is None or min(dates) < best):
                        best = min(dates)
            gates[dist.upper()] = best
            print(f"  {st_name}: {dist} labor counted from "
                  f"{best if best else 'never (no debriefs yet)'}")
        out[st_name] = gates
    return out


# Sentinel distinguishing "station has no launch gate" from "gated with no
# debriefs yet" (a None date suppresses ALL labor).
_NO_LAUNCH_GATE = object()


def worked_hours(cfg, hsel_shift, hsel_plain, emp, year, month, ndays,
                 gates=None, launch=_NO_LAUNCH_GATE):
    """Per-day worked hours for one station config in one month: hourly
    punches + open-shift estimates + 40/7 per active salaried head, with
    optional per-dist gates and a station-wide launch gate.

    This is the single calculation contract for worked hours — shared by the
    commercial build (build_month) and build_mro_hours.py. Change it here and
    both pulses move together.

    cfg keys used: labor_keys, salary_keys (explicit [] = no salaried
    imputation), facility (True = plain calendar-day attribution),
    shift_window. Returns (hourly, inc_n, sal_counts, worked), lists indexed
    by day-1. inc_n counts punches still open when the hours file was cut:
    they contribute no hours (the pulse does not estimate) and only drive the
    asterisk on the page.
    """
    days = list(range(1, ndays + 1))
    gates = gates or {}
    keys = [k.upper() for k in (cfg.get("labor_keys") or [cfg["labor_key"]]) if k]
    if isinstance(cfg.get("salary_keys"), list):
        # an explicit [] means this crew carries no salaried imputation
        skeys = [k.upper() for k in cfg["salary_keys"] if k]
    else:
        skeys = keys
    hourly = [0.0] * ndays
    inc_n = [0] * ndays
    # A crew whose window wraps midnight works overnight, so it takes the
    # shift-back attribution; a daytime crew takes the plain calendar day.
    win = cfg.get("shift_window")
    if win:
        overnight = win["from"] > win["to"]
        hsel = hsel_shift if overnight else hsel_plain
    else:
        hsel = hsel_plain if cfg.get("facility") else hsel_shift
    if win:
        lo, hi = win["from"], win["to"]
        inwin = ((hsel["punch_hour"] >= lo) | (hsel["punch_hour"] < hi)) \
            if overnight else \
            ((hsel["punch_hour"] >= lo) & (hsel["punch_hour"] < hi))
        hsel = hsel[inwin & (hsel["punch_hour"] >= 0)]
    # Per-dist override: a DAY crew pooled into an overnight station keeps
    # plain calendar-day attribution (Sam, 2026-09-28 — DFW WIDEBODY never
    # works past midnight; shift-back would push its morning punches onto
    # yesterday). ASK when adding any new dist: does the crew work past
    # midnight? See CLAUDE.md.
    plain_dists = {k.upper() for k in (cfg.get("plain_day_dists") or [])}

    def gated_out(dist, dnum):
        g = gates.get(dist)
        return dist in gates and (g is None or date(year, month, dnum) < g)

    hourly_pool = hsel[hsel["pay_type"] == "Hourly"]
    plain_pool = hsel_plain[hsel_plain["pay_type"] == "Hourly"]
    for key in keys:
        pool = plain_pool if key in plain_dists else hourly_pool
        sub = pool[pool["labor_dist"].str.upper() == key]
        if sub.empty:
            continue
        for d, v in sub.groupby("day")["hours"].sum().items():
            if 1 <= d <= ndays and not gated_out(key, d):
                hourly[d - 1] += float(v)
        openrows = sub[sub["incomplete"] == 1]
        for d, n in openrows.groupby("day").size().items():
            if 1 <= d <= ndays and not gated_out(key, d):
                inc_n[d - 1] += int(n)

    sal_counts = [0] * ndays
    for skey in skeys:
        sal_emp = emp[(emp["pay_type"] == "Salary")
                      & (emp["labor_dist"].str.upper() == skey)]
        if sal_emp.empty:
            continue
        for i, dnum in enumerate(days):
            if gated_out(skey, dnum):
                continue
            cut = date(year, month, dnum)
            sal_counts[i] += sum(
                1 for ld, sf in zip(sal_emp["last_day"], sal_emp["salary_from"])
                if ld and ld > cut and (sf is None or cut >= sf))
    # Pre-launch labor (before this station's first debrief) is
    # implementation/training — zero it out rather than score it.
    if launch is not _NO_LAUNCH_GATE:
        for i, dnum in enumerate(days):
            if launch is None or date(year, month, dnum) < launch:
                hourly[i] = 0.0
                inc_n[i] = 0
                sal_counts[i] = 0

    worked = [round(h + n * 40 / 7, 2) for h, n in zip(hourly, sal_counts)]
    return hourly, inc_n, sal_counts, worked


def build_month(year, month, stations, tables, hours, emp, closeout, hours_start,
                labor_gates):
    ndays = month_days(year, month)
    is_current = (year, month) == (TODAY.year, TODAY.month)
    is_past = date(year, month, 1) < date(TODAY.year, TODAY.month, 1)
    # elapsed-day cutoff, exactly as the sheets do it: past month => all days,
    # current month => days strictly before today
    cutoff = ndays + 1 if is_past else (TODAY.day if is_current else 0)

    # Days the AA feed simply has not delivered, and the open Closeout
    # Compare rows landing in this month — both drive asterisks, neither
    # invents a number.
    last_elapsed = ndays if is_past else max(cutoff - 1, 0)
    aa_gaps = aa_missing_days(tables.get("AA_Debriefs"), year, month,
                              last_elapsed)
    co_month = (closeout[closeout["date"].map(
        lambda d: d is not None and d.year == year and d.month == month)]
        if len(closeout) else closeout)

    # pre-slice each debrief table to this month. NB: a zero-row table must
    # keep its columns — `df[[]]` is COLUMN selection in pandas, which
    # silently drops them (first hit by the brand-new Widebody table,
    # 2026-09-28), so empty tables short-circuit through head(0).
    month_tbl = {}
    for tname, df in tables.items():
        mask = [d is not None and d.year == year and d.month == month
                for d in df["date"]]
        sel = df[mask].copy() if mask else df.head(0).copy()
        sel["day"] = [d.day for d in sel["date"]]
        month_tbl[tname] = sel

    def month_slice(col):
        sel = hours[[d.year == year and d.month == month for d in hours[col]]].copy()
        sel["day"] = [d.day for d in sel[col]]
        return sel

    hsel_shift = month_slice("attr_date")        # commercial stations
    hsel_plain = month_slice("attr_date_plain")  # facility stations

    out = {}
    for st_name, cfg in stations.items():
        days = list(range(1, ndays + 1))
        code = st_name.strip()[:3].upper()

        svc_rows = []
        budgeted = [0.0] * ndays
        for svc in cfg["services"]:
            vals = []
            # Does this service come from the AA feed? Only those days can
            # be blank purely because AA has not delivered.
            uses_aa = any(sp["table"] == "AA_Debriefs"
                          for sp in svc.get("specs", []))
            starts = service_start(svc)
            if svc["kind"] == "fixed":
                daily = dow_daily_rates(svc)
                for d in days:
                    elapsed_day = is_past or d < cutoff
                    if not elapsed_day or (starts and
                                           date(year, month, d) < starts):
                        vals.append(0)
                        continue
                    v = (daily[date(year, month, d).weekday()] if daily
                         else svc["rate"])
                    vals.append(round(v or 0, 2))
            else:
                for d in days:
                    if starts and date(year, month, d) < starts:
                        vals.append(0)
                        continue
                    total = 0
                    for spec in svc["specs"]:
                        t = month_tbl.get(spec["table"])
                        if t is None or t.empty:
                            continue
                        m = spec_mask(t, spec) & (t["day"] == d)
                        if spec["fn"] == "SUMIFS":
                            total += float(t.loc[m, spec["sum_col"]].sum())
                        else:
                            total += int(m.sum())
                    vals.append(total)
            rate = svc.get("rate") or 0
            # Hours this service contributes, per day. Three ways to earn a
            # budget: a flat daily allowance, a rate per aircraft serviced,
            # and — where the work is priced per airframe rather than timed
            # (NetJets at FLL) — the service's own REVENUE at a fixed
            # dollars-per-budgeted-hour (Sam, 2026-10-07: $45 = 1 h).
            #
            # `hours_from` splits what a row SHOWS from what it EARNS: the
            # row still counts aircraft serviced, but its hours come from the
            # money column beside the count. Without it a NetJets row would
            # need one hours-per-job rate, and there isn't one — a Complete
            # Exterior is $331 on an EMB-505S and $1,071 on a GL7500.
            money = svc.get("hours_from")
            if money:
                hrs = []
                for d in days:
                    if starts and date(year, month, d) < starts:
                        hrs.append(0.0)
                        continue
                    tot = 0.0
                    for spec in money:
                        t = month_tbl.get(spec["table"])
                        if t is None or t.empty:
                            continue
                        m = spec_mask(t, spec) & (t["day"] == d)
                        tot += float(t.loc[m, spec["sum_col"]].sum())
                    hrs.append(round(tot / rate, 2) if rate else 0.0)
            elif svc["kind"] == "fixed":
                hrs = [round(v, 2) for v in vals]
            elif svc["kind"] == "revenue":
                hrs = [round(v / rate, 2) if rate else 0.0 for v in vals]
            else:
                hrs = [round(v * rate, 2) for v in vals]
            for i, h in enumerate(hrs):
                budgeted[i] += h
            svc_rows.append({"name": svc["name"], "kind": svc["kind"],
                             "rate": rate, "days": vals, "hours": hrs,
                             "aircraft": svc.get("aircraft", False),
                             "uses_aa": uses_aa,
                             "group": svc.get("group"),
                             "priced": bool(svc.get("hours_from")),
                             "hidden": bool(svc.get("hidden"))})

        # Collapsible groups (Sam, 2026-10-07, for FLL): a station may sort
        # its services under named groups. The GROUP row carries the budget
        # HOURS its members earn; the member rows stay what they have always
        # been — counts of aircraft serviced — and are hidden until opened.
        # A service with no group keeps rendering flat, so every other
        # station is untouched. A `hidden` member earns hours without a row
        # of its own: NetJets' budget comes from revenue, which is not one
        # of the services anybody counts.
        groups = []
        if any(r["group"] for r in svc_rows):
            for g in cfg.get("group_order") or []:
                members = [r for r in svc_rows if r["group"] == g]
                if not members:
                    continue
                hrs = [round(sum(m["hours"][i] for m in members), 2)
                       for i in range(ndays)]
                groups.append({"name": g, "hours": hrs,
                               "children": [m["name"] for m in members
                                            if not m["hidden"]]})
            missing = sorted({r["group"] for r in svc_rows if r["group"]}
                             - {g["name"] for g in groups})
            if missing:
                print(f"  !! {st_name}: service group(s) {missing} not in "
                      f"group_order — they will not render")

        # worked hours: hourly punches + salaried imputation, per day —
        # the shared contract in worked_hours()
        hourly, inc_n, sal_counts, worked = worked_hours(
            cfg, hsel_shift, hsel_plain, emp, year, month, ndays,
            gates=labor_gates.get(st_name, {}),
            launch=hours_start.get(st_name, _NO_LAUNCH_GATE))

        # weekly blocks + stats (avg over elapsed days only)
        weeks = []
        for w in range(4):
            lo = w * 7 + 1
            hi = min(lo + 6, ndays) if w < 3 else ndays
            drange = list(range(lo, hi + 1))
            elapsed = [d for d in drange if is_past or d < cutoff]
            def avg(series):
                pts = [series[d - 1] for d in elapsed]
                return round(sum(pts) / len(pts), 2) if pts else None
            ac = None
            if not cfg["fac_only"]:
                acc = 0.0
                got = False
                for sr in svc_rows:
                    if not sr["aircraft"]:
                        continue
                    a = avg(sr["days"])
                    if a is not None:
                        acc += a
                        got = True
                ac = round(acc, 2) if got else None
            ab, aw = avg(budgeted), avg(worked)
            if ab and aw is not None and ab != 0:
                ratio = aw / ab
                pct = round(abs(1 - ratio), 2) * 100
                var = {"pct": round(pct), "dir": "Over Budget" if ratio > 1 else "Under Budget"}
            else:
                var = None
            weeks.append({"days": drange, "avg_aircraft": ac,
                          "avg_budgeted": ab, "avg_worked": aw, "variance": var})

        # ---- asterisk notes: what the numbers on this row cannot show
        worked_notes, budget_notes = {}, {}
        for i, dnum in enumerate(days):
            if inc_n[i]:
                n = inc_n[i]
                worked_notes[str(dnum)] = (
                    f"{n} shift{'' if n == 1 else 's'} had not been punched "
                    f"out when the hours file was cut, so {'its' if n == 1 else 'their'} "
                    f"hours are not included here. The figure is the hours "
                    f"actually recorded, and will rise once the punches close.")
        st_aa_gaps = aa_gaps.get(code, set()) if any(
            sr.get("uses_aa") for sr in svc_rows) else set()
        st_tables = {sp["table"] for sv in cfg["services"]
                     for sp in sv.get("specs", [])}
        # {day: [missing, double]} for this station. Built with vectorized
        # masks — indexing a frame with an empty list selects columns, not
        # rows, which silently drops the schema.
        co_day = {}
        if len(co_month):
            sel = co_month[co_month["loc"].eq(code)
                           & (co_month["table"].isin(st_tables)
                              | co_month["table"].isna())]
            for _, cr in sel.iterrows():
                slot = co_day.setdefault(cr["date"].day, [0, 0])
                slot[0 if cr["kind"] == "missing" else 1] += 1
        for dnum in days:
            notes = []
            if dnum in st_aa_gaps:
                notes.append("AA has not delivered debriefs for this day, so "
                             "its AA job counts are missing rather than zero.")
            if dnum in co_day:
                miss, dub = co_day[dnum]
                if miss:
                    notes.append(
                        f"Closeout Compare has {miss} job{'' if miss == 1 else 's'} "
                        f"on the closeout with no debrief submitted. Those are "
                        f"NOT added here — the count shows only debriefed work, "
                        f"so it is understated until the debrief is filed.")
                if dub:
                    notes.append(
                        f"Closeout Compare found {dub} duplicate "
                        f"debrief{'' if dub == 1 else 's'}; the extra "
                        f"submission{'' if dub == 1 else 's'} "
                        f"{'has' if dub == 1 else 'have'} been removed from "
                        f"this count.")
            if notes:
                budget_notes[str(dnum)] = " ".join(notes)

        mtd_days = [d for d in days if is_past or d < cutoff]
        mtd_b = round(sum(budgeted[d - 1] for d in mtd_days), 1)
        mtd_w = round(sum(worked[d - 1] for d in mtd_days), 1)
        out[st_name] = {
            # Which vertical this location is, so the page can offer the
            # scheduler link on a Private/MRO location (Sam, 2026-10-11).
            "vertical": cfg.get("vertical", "commercial"),
            "services": svc_rows,
            "groups": groups,
            "budgeted": [round(b, 2) for b in budgeted],
            "worked": worked,
            "hourly": [round(h, 2) for h in hourly],
            "incomplete_n": inc_n,
            "worked_notes": worked_notes,
            "budget_notes": budget_notes,
            "salary_heads": sal_counts,
            "weeks": weeks,
            "mtd": {"budgeted": mtd_b, "worked": mtd_w,
                    "variance": (round(abs(1 - mtd_w / mtd_b) * 100)
                                 if mtd_b else None),
                    "over": mtd_w > mtd_b if mtd_b else None},
            "elapsed_through": (cutoff - 1) if not is_past else ndays,
        }
    return {"year": year, "month": month, "ndays": ndays,
            "label": date(year, month, 1).strftime("%B %Y"), "stations": out}


def check_sources_present():
    """Fail fast and by name if a source the build needs wasn't fetched.
    Guards against the fetch list and the build drifting apart."""
    from sources import SOURCE_NAMES
    roots = {"This Years Hours.csv": PAYLOCITY, "Basic Employee Info.csv": PAYLOCITY,
             "Location Management.csv": LOCMGMT_DIR, "Service Budgets.xlsx": BUDGETS_DIR,
             "Private and MRO Pulse Locations.xlsx": BUDGETS_DIR,
             "Early Terminations.csv": DEFINITIVE_DIR,
             "Pay Type Changes.csv": DEFINITIVE_DIR}
    from sources import SOURCE_FOLDERS
    missing = [n for n in SOURCE_NAMES if not (roots.get(n, DEBRIEFS) / n).exists()]
    # A mirrored folder that arrived empty is worse than a missing file: the
    # build would succeed and publish Private/MRO locations at a 0 budget
    # against real worked hours (2026-10-10). Fail where it can be seen.
    for folder in SOURCE_FOLDERS:
        d = BUDGETS_DIR / folder.rsplit("/", 1)[-1]
        if not d.is_dir() or not any(d.glob("*.json")):
            missing.append(folder + "/ (no documents)")
    if missing:
        raise SystemExit(
            "missing source file(s): " + ", ".join(missing) +
            "\nIf this is CI, check that sources.py lists them and "
            "fetch_sources.py downloaded them.")


def main():
    check_sources_present()
    base_stations = json.loads((HERE / "stations.json").read_text())
    catalog = json.loads((HERE / "service_catalog.json").read_text())
    catalog.update(json.loads((HERE / "catalog_extras.json").read_text()))
    overrides = json.loads((HERE / "station_overrides.json").read_text())
    budgets = load_budget_workbook()
    print(f"Service Budgets.xlsx: {len(budgets)} location sheets")
    stations = merge_budget_config(base_stations, budgets, catalog, overrides)
    # Private/MRO locations join as ordinary stations (Sam, 2026-10-11) — one
    # list, every location, whichever way its budget is earned.
    mro_cfgs = _mro_configs()
    mro_jobs = load_mro_jobs(mro_cfgs)
    clash = sorted(set(mro_cfgs) & set(stations))
    if clash:
        raise SystemExit(f"location name used by both a Service Budgets sheet "
                         f"and an MRO sheet: {clash}")
    stations.update(mro_stations(mro_cfgs, mro_jobs))
    stations = dict(sorted(stations.items(), key=lambda kv: kv[0].upper()))
    print(f"Private/MRO: {len(mro_cfgs)} location(s); {len(stations)} in total")
    print("loading sources...")
    tables = {
        "Envoy_Debriefs": load_envoy(),
        "GoJet_Debriefs": load_gojet(),
        "Mesa_Debriefs": load_mesa(),
        "PSA_Debriefs": load_psa(),
        "Breeze_Debriefs": load_breeze(),
        "Ultra_Debriefs": load_ultra(),
        "Widebody_Debriefs": load_widebody(),
        "Frontier_Debriefs": load_frontier(),
        "AA_Debriefs": load_aa(),
        "APU_Wash": load_apu(),
        "NetJets_Debriefs": load_netjets(),
        "JSX_Debriefs": load_jsx(),
        "MRO_Jobs": mro_jobs,
    }
    for k, v in tables.items():
        print(f"  {k}: {len(v)} rows")
    emp = load_employees()
    hours = load_hours(emp)
    print(f"  employees: {len(emp)}, punch rows in window: {len(hours)}")

    closeout = load_closeout()
    apply_double_debriefs(tables, closeout)
    hours_start = first_debrief_dates(stations, tables)
    labor_gates = labor_gate_dates(stations, tables)
    months = []
    y, m = WINDOW_START.year, WINDOW_START.month
    while (y, m) <= (TODAY.year, TODAY.month):
        months.append(build_month(y, m, stations, tables, hours, emp, closeout,
                                  hours_start, labor_gates))
        m += 1
        if m == 13:
            y, m = y + 1, 1
    data = {
        "generated": datetime.now().strftime("%b %d, %Y %I:%M %p"),
        "months": months,
        "station_names": list(stations.keys()),
        "managers": load_managers(stations.keys()),
    }
    (HERE / "pulse_data.json").write_text(json.dumps(data))
    print(f"wrote pulse_data.json ({(HERE / 'pulse_data.json').stat().st_size // 1024} KB, "
          f"{len(months)} months)")

    template = (HERE / "template.html").read_text(encoding="utf-8")
    html = template.replace("/*__DATA__*/", json.dumps(data))
    (HERE / "index.html").write_text(html, encoding="utf-8")
    print(f"wrote index.html ({(HERE / 'index.html').stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
