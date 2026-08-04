# foxtrot-labor-pulse

**Owner:** Samuel Kosco — Data Analyst, Foxtrot Aviation Services
**Purpose:** Read-only Labor Pulse dashboard for ops — the web replacement for opening `Live Commercial Pulse Sheet.xlsx`. Budgeted vs worked labor hours per station per day, with weekly averages and variance.

The page is intentionally uneditable: ops get a link, not the workbook. The repo name carries a random suffix and the page carries `noindex` because GitHub Pages is public — treat the URL as the access control and share it only internally.

## How it works

1. `.github/workflows/refresh.yml` runs daily (dual cron 13:00/14:00 UTC; a guard step lets only the 9 AM America/New_York run proceed, handling DST) plus manual dispatch.
2. `fetch_sources.py` downloads the 13 source files from the Data Hub SharePoint drive via Microsoft Graph (client credentials — same Foxtrot Report Automation app and drive ID as the compliance trackers). Beyond the original pulse-workbook sources this includes `AA Debriefs.csv` (STL/PVD AA cabin jobs: Job Type DTC/RSTC→AA Turn, RON/RRON→AA RON, Security→AA Security, Ultra→Ultra, Shroud→Shroud Cleaning; Status=Completed only) and `APU Wash.xlsx` Sheet2 (JFK/MCO APU washes by Location, Status=Completed only).
3. `build_pulse.py` reproduces the workbook's full calculation chain from those sources (the workbook itself is never read) and renders `index.html` from `template.html` with the data embedded.
4. The workflow commits `index.html` + `pulse_data.json`; GitHub Pages serves it.

## Calculation notes

- **Worked Hours** = hourly punched hours (day attributed by punch-in time minus 12 h, so overnight shifts land on the day the shift started; "Not Defined" labor dist repaired from the employee master) + 40/7 h per active salaried employee in the station's labor dist. Directors/Regional Managers are excluded, matching the workbook query.
- **Budgeted Hours** = per-service aircraft counts (from the airline debrief workbooks, Yes→1 flag logic identical to the Power Queries) × the station's hours-per-job rates, plus flat daily facility budgets released only for elapsed days.
- Weekly averages cover elapsed days only; variance = |1 − worked/budgeted| rounded like the sheet.
- **`Pulse Sheets/Service Budgets.xlsx` is the owner-editable control panel.** One sheet per location, columns Service | Budget. It decides which locations the pulse shows, which services each has, and the hours-per-job rates — edits flow into the page at the next refresh. Adding a sheet adds a location: the airport code is the first 3 letters of the sheet name; the labor dist defaults to `<CODE> FAC` if the sheet name contains "FAC", else `<CODE> CABIN`; service rows are interpreted via `service_catalog.json` (recognized names like "Envoy IHC", "PSA RON", "GoJet CE", "Facility Budget" — matching is case/punctuation-insensitive). Unknown service names are treated as a flat daily budget of that many hours, with a build-log warning.
- `stations.json` is the per-station spec config **auto-derived from the pulse workbook's formulas** by `extract_config.py` (local-only script). It still supplies the counting rules and labor-dist exceptions (TYS PSA, STL GOJET, IAH CABIN+FAC) for existing stations; `make_budget_workbook.py` regenerates Service Budgets.xlsx and `service_catalog.json` from it.
- `catalog_extras.json` holds hand-added service templates (AA Turn/RON/Security, Ultra, APU Wash) merged over the generated catalog — regenerating `service_catalog.json` won't lose them. `station_overrides.json` holds per-station exceptions: STL AA CAB / PVD AA CAB labor dist codes match their sheet names (not the `<CODE> CABIN` convention), and STL AA CAB's "Shroud Cleaning" counts from AA_Debriefs, not the Ultra workbook DFW uses.
- The build intentionally **fixes four workbook bugs**: CMH counting LIT's Envoy aircraft (days 2–31), IAH budget rates shifted one row, BDL week-4 hardcoded rates, FLL/TUS FAC days 29–31 pulling MLB FAC's worked hours. Expect those cells to differ from Excel until the workbook is fixed.
- Data window starts 2026-05-01 (`WINDOW_START`), mirroring the hardcoded cutoff in the workbook's M queries.
- **Master page** ("All Locations — Overview", the default landing view): company-wide MTD KPIs, a per-station table, and By Regional Manager groups. Manager mapping comes from `Power BI Data Sources/Location Management.csv`, matched **blindly on the first 3 letters** of the location on both sides (the airport code) — per the owner's instruction. A location with multiple managers in the CSV appears under each of them, so manager totals overlap and do not sum to the company total.

## Secrets

Same three as `envoy-compliance-tracker` (see that repo's CLAUDE.md): `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` (Foxtrot Report Automation app — secret expires every 24 months; renew in Entra and update here too).

## Local development

`python build_pulse.py` with no env vars reads the synced Data Hub folder directly. `PULSE_DATA_DIR=sources python build_pulse.py` after `python fetch_sources.py` mimics CI. Open `index.html` — no server needed.
