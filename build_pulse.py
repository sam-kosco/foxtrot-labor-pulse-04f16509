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
    DEBRIEFS = PAYLOCITY = LOCMGMT_DIR = BUDGETS_DIR = Path(os.environ["PULSE_DATA_DIR"])
else:
    DH = Path(r"C:\Users\samko\Foxtrot Aviation Services\Data Hub - Documents")
    DEBRIEFS = DH / "Power Flows" / "Debriefs"
    PAYLOCITY = DH / "Paylocity Reports"
    LOCMGMT_DIR = DH / "Power BI Data Sources"
    BUDGETS_DIR = DH / "Pulse Sheets"
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
    out = pd.DataFrame({"date": t["Date"].map(to_date)})
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
        "Breeze RON": pd.to_numeric(t["Breeze RON"], errors="coerce").fillna(0).astype(int),
        "Breeze Ultra": pd.to_numeric(t["Breeze Ultra"], errors="coerce").fillna(0).astype(int),
    })


def load_ultra():
    t = read_table(DEBRIEFS / "Ultra Debriefs.xlsx", "Table1")
    out = pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Location"].astype(str).str.strip(),
    })
    shroud = next((c for c in t.columns if c and "Shroud" in str(c)), None)
    out["Shroud Count"] = (
        pd.to_numeric(t[shroud], errors="coerce").fillna(0) if shroud else 0
    )
    return out


def load_frontier():
    t = read_table(DEBRIEFS / "Frontier Debriefs.xlsx", "Table1")
    t = t[t["Status"] == "Complete"]
    return pd.DataFrame({
        "date": t["Date"].map(to_date),
        "Location": t["Location"].astype(str).str.strip(),
        "Aircraft Type and Service": t["Aircraft Type and Service"],
    })


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


def merge_budget_config(stations, budgets, catalog):
    """Build the effective per-station config: the budget workbook decides the
    station list, service list, and rates; specs come from stations.json when
    the station+service already exist there, else from the service catalog
    templates ({LOC} = first 3 letters of the sheet name). Unknown service
    names fall back to a flat daily budget, with a warning."""
    merged = {}
    for st_name, rows in budgets.items():
        code = st_name.strip()[:3].upper()
        base = stations.get(st_name)
        base_by_name = ({NORM(s["name"]): (i, s) for i, s in enumerate(base["services"])}
                        if base else {})
        base_aircraft = ({r - 3 for r in base["aircraft_service_rows"]} if base else set())
        services = []
        for svc_name, rate in rows:
            key = NORM(svc_name)
            if key in base_by_name:
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
        if base:
            labor_keys, salary_keys = base["labor_keys"], base["salary_keys"]
        else:
            k = f"{code} FAC" if "FAC" in st_name.upper() else f"{code} CABIN"
            labor_keys = salary_keys = [k]
            print(f"  new location {st_name!r}: labor dist {k!r} (convention)")
        merged[st_name] = {"labor_keys": labor_keys, "salary_keys": salary_keys,
                           "fac_only": fac_only, "services": services}
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


def load_employees():
    e = pd.read_csv(PAYLOCITY / "Basic Employee Info.csv", dtype=str, encoding="cp1252")
    e = e[~e["Job Title"].isin(EXCLUDED_TITLES)].copy()
    term = pd.to_datetime(e["Termination Date"], errors="coerce")
    active = e["Employee Status Description"].str.strip() == "Active"
    e["last_day"] = [
        TODAY if a else (t.date() if pd.notna(t) else None)
        for a, t in zip(active, term)
    ]
    e["labor_dist"] = e["Labor Dist Description"].fillna("").str.strip()
    e["pay_type"] = e["Pay Type Code"].fillna("").str.strip()
    e["id"] = e["Employee Id"].str.strip()
    return e


def load_hours(emp):
    h = pd.read_csv(PAYLOCITY / "This Years Hours.csv", dtype=str, encoding="cp1252")
    h["work_date"] = pd.to_datetime(h["Work Date"], errors="coerce").dt.date
    h = h[(h["work_date"] >= WINDOW_START) & (h["work_date"] <= TODAY)].copy()
    h["hours"] = pd.to_numeric(h["Summation of Paid Duration (hours)"], errors="coerce").fillna(0)
    punch = pd.to_datetime(h["Punch In Time"], errors="coerce")
    # workbook helper: DAY/MONTH(punch-in − 12h) — overnight shifts credit the
    # day the shift started; rows with no punch-in fall back to Work Date
    shifted = punch - timedelta(hours=12)
    h["attr_date"] = [
        s.date() if pd.notna(s) else w for s, w in zip(shifted, h["work_date"])
    ]
    id2dist = dict(zip(emp["id"], emp["labor_dist"]))
    id2pay = dict(zip(emp["id"], emp["pay_type"]))
    eid = h["Employee Id"].str.strip()
    raw = h["Labor Dist Name"].fillna("").str.strip()
    h["labor_dist"] = [
        id2dist.get(i, "") if r == "Not Defined" else r for i, r in zip(eid, raw)
    ]
    h["pay_type"] = eid.map(id2pay).fillna("")
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


def build_month(year, month, stations, tables, hours, emp):
    ndays = month_days(year, month)
    is_current = (year, month) == (TODAY.year, TODAY.month)
    is_past = date(year, month, 1) < date(TODAY.year, TODAY.month, 1)
    # elapsed-day cutoff, exactly as the sheets do it: past month => all days,
    # current month => days strictly before today
    cutoff = ndays + 1 if is_past else (TODAY.day if is_current else 0)

    # pre-slice each debrief table to this month
    month_tbl = {}
    for tname, df in tables.items():
        sel = df[[d is not None and d.year == year and d.month == month
                  for d in df["date"]]].copy()
        sel["day"] = [d.day for d in sel["date"]]
        month_tbl[tname] = sel

    hsel = hours[[d.year == year and d.month == month for d in hours["attr_date"]]].copy()
    hsel["day"] = [d.day for d in hsel["attr_date"]]

    out = {}
    for st_name, cfg in stations.items():
        days = list(range(1, ndays + 1))
        svc_rows = []
        budgeted = [0.0] * ndays
        for svc in cfg["services"]:
            vals = []
            if svc["kind"] == "fixed":
                for d in days:
                    v = svc["rate"] if (is_past or d < cutoff) else 0
                    vals.append(round(v or 0, 2))
            else:
                for d in days:
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
            for i, v in enumerate(vals):
                budgeted[i] += (v if svc["kind"] == "fixed" else v * rate)
            svc_rows.append({"name": svc["name"], "kind": svc["kind"],
                             "rate": rate, "days": vals,
                             "aircraft": svc.get("aircraft", False)})

        # worked hours: hourly punches + salaried imputation, per day
        keys = [k.upper() for k in (cfg.get("labor_keys") or [cfg["labor_key"]]) if k]
        skeys = [k.upper() for k in (cfg.get("salary_keys") or keys) if k] or keys
        hourly = [0.0] * ndays
        st_h = hsel[(hsel["labor_dist"].str.upper().isin(keys))
                    & (hsel["pay_type"] == "Hourly")]
        for d, v in st_h.groupby("day")["hours"].sum().items():
            if 1 <= d <= ndays:
                hourly[d - 1] = float(v)
        sal_emp = emp[(emp["pay_type"] == "Salary")
                      & (emp["labor_dist"].str.upper().isin(skeys))]
        sal_counts = []
        for d in days:
            cut = date(year, month, d)
            n = sum(1 for ld in sal_emp["last_day"] if ld and ld > cut)
            sal_counts.append(n)
        worked = [round(h + n * 40 / 7, 2) for h, n in zip(hourly, sal_counts)]

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

        mtd_days = [d for d in days if is_past or d < cutoff]
        mtd_b = round(sum(budgeted[d - 1] for d in mtd_days), 1)
        mtd_w = round(sum(worked[d - 1] for d in mtd_days), 1)
        out[st_name] = {
            "services": svc_rows,
            "budgeted": [round(b, 2) for b in budgeted],
            "worked": worked,
            "hourly": [round(h, 2) for h in hourly],
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


def main():
    base_stations = json.loads((HERE / "stations.json").read_text())
    catalog = json.loads((HERE / "service_catalog.json").read_text())
    budgets = load_budget_workbook()
    print(f"Service Budgets.xlsx: {len(budgets)} location sheets")
    stations = merge_budget_config(base_stations, budgets, catalog)
    print("loading sources...")
    tables = {
        "Envoy_Debriefs": load_envoy(),
        "GoJet_Debriefs": load_gojet(),
        "Mesa_Debriefs": load_mesa(),
        "PSA_Debriefs": load_psa(),
        "Breeze_Debriefs": load_breeze(),
        "Ultra_Debriefs": load_ultra(),
        "Frontier_Debriefs": load_frontier(),
    }
    for k, v in tables.items():
        print(f"  {k}: {len(v)} rows")
    emp = load_employees()
    hours = load_hours(emp)
    print(f"  employees: {len(emp)}, punch rows in window: {len(hours)}")

    months = []
    y, m = WINDOW_START.year, WINDOW_START.month
    while (y, m) <= (TODAY.year, TODAY.month):
        months.append(build_month(y, m, stations, tables, hours, emp))
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
