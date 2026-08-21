# foxtrot-labor-pulse

**Owner:** Samuel Kosco — Data Analyst, Foxtrot Aviation Services
**Purpose:** Read-only Labor Pulse dashboard for ops — the web replacement for opening `Live Commercial Pulse Sheet.xlsx`. Budgeted vs worked labor hours per station per day, with weekly averages and variance.

The page is intentionally uneditable: ops get a link, not the workbook. The repo name carries a random suffix and the page carries `noindex` because the Pages site is public — treat the URL as the access control and share it only internally. Since the org migration the CODE is private (org repo); only the rendered site is public. The platform embeds this site and its JSON feeds the platform homepage stats.

## How it works

1. `.github/workflows/refresh.yml` runs daily (dual cron 13:00/14:00 UTC; a guard step lets only the 9 AM America/New_York run proceed, handling DST) plus manual dispatch.
2. `fetch_sources.py` downloads the 13 source files from the Data Hub SharePoint drive via Microsoft Graph (client credentials — same Foxtrot Report Automation app and drive ID as the compliance trackers). Beyond the original pulse-workbook sources this includes `AA Debriefs.csv` (STL/PVD AA cabin jobs: Job Type DTC/RSTC→AA Turn, RON/RRON→AA RON, Security→AA Security, Ultra→Ultra, Shroud→Shroud Cleaning; Status=Completed only) and `APU Wash.xlsx` Sheet2 (JFK/MCO). **APU Wash Sheet2 dropped its Status column in Aug 2026** in favour of one numeric column per service — `APU Wash`, `Landing Gear Bay Wash`, `Partial Belly Wash` — valued 1 (completed), 0.5 (cancelled, half billable) or 0 (no job). Only a 1 counts toward the pulse; `load_apu` raises if any of the three columns goes missing rather than silently counting zero.
3. `build_pulse.py` reproduces the workbook's full calculation chain from those sources (the workbook itself is never read) and renders `index.html` from `template.html` with the data embedded.
4. The workflow commits `index.html` + `pulse_data.json`; GitHub Pages serves it.

## Calculation notes

- **Worked Hours** = hourly punched hours + 40/7 h per active salaried employee in the station's labor dist ("Not Defined" labor dist repaired from the employee master). Directors/Regional Managers are excluded, matching the workbook query.
- **Day attribution differs by station type.** Commercial stations use the workbook's rule — punch-in minus 12 h, so an overnight shift credits the day it started. **Facility stations use the plain calendar day of the punch** (owner decision, Aug 2026): they don't run overnight shifts, and the 12 h shift-back was pushing normal morning starts (TUS FAC begins 5–6 AM) onto the previous day. A station counts as facility when its sheet name ends in "FAC" or all its services are flat daily budgets (`facility` flag, set in `merge_budget_config`). Note `IAH` draws on both `IAH CABIN` and `IAH FAC` but is a commercial sheet, so all its hours keep the shift-back.
- **Holiday/PTO pay is excluded** (rows with a punch-in, no punch-out, and paid hours > 0 — owner decision Aug 2026; the Excel workbook still counts these).
- **AA debrief gaps are filled**: AA's feed runs through AA's internal system and doesn't update automatically, so past dates can be wholly absent per station. Any past date with zero AA rows gets the per-service 7-day average taken from the 7 days ending at the most recent listed date before it (`aa_fill_plan`). Filled cells show `~` on the page and real data replaces them when the feed delivers.
- The page force-reloads itself every 30 minutes so an open tab can't show stale data.
- **Open shifts are estimated**: rows with a punch-in, no punch-out, and 0 paid hours whose punch-in falls in the 18 h before This Years Hours.csv was last modified (SharePoint lastModifiedDateTime captured in sources_meta.json by fetch_sources.py; local file mtime as fallback) count as the employee's median closed-shift duration (fallback location median, then 8 h; capped 16 h). Estimated cells show a `~` and the tooltip breaks the estimate out; real punches replace estimates at the next refresh. Future/scheduled placeholder punches and stale opens fall outside the window and count 0.
- **Budgeted Hours** = per-service aircraft counts (from the airline debrief workbooks, Yes→1 flag logic identical to the Power Queries) × the station's hours-per-job rates, plus flat daily facility budgets released only for elapsed days.
- Weekly averages cover elapsed days only; variance = |1 − worked/budgeted| rounded like the sheet.
- **`Pulse Sheets/Service Budgets.xlsx` is the owner-editable control panel.** One sheet per location, columns Service | Budget. It decides which locations the pulse shows, which services each has, and the hours-per-job rates — edits flow into the page at the next refresh. Adding a sheet adds a location: the airport code is the first 3 letters of the sheet name; the labor dist defaults to `<CODE> FAC` if the sheet name contains "FAC", else `<CODE> CABIN`; service rows are interpreted via `service_catalog.json` (recognized names like "Envoy IHC", "PSA RON", "GoJet CE", "Facility Budget" — matching is case/punctuation-insensitive). Unknown service names are treated as a flat daily budget of that many hours, with a build-log warning.
- `stations.json` is the per-station spec config **auto-derived from the pulse workbook's formulas** by `extract_config.py` (local-only script). It still supplies the counting rules and labor-dist exceptions (TYS PSA, STL GOJET, IAH CABIN+FAC) for existing stations; `make_budget_workbook.py` regenerates Service Budgets.xlsx and `service_catalog.json` from it.
- **JSX stations** (`BUR-SNA`, `TEB-HPN`, `DAL`, `SCF`) read `Power Flows/Debriefs/JSX Debriefs.xlsx`, table `Debriefs` on Sheet1 — one row per tail with 0/1 columns for RON, Interior Detail, Exterior Detail, Carpet Extraction, keyed by `Service Location`. Two sheets each cover two airports (Excel forbids `/` in sheet names, hence the hyphen); their specs list one COUNTIFS per airport and the build sums them. **Rows dated before `JSX_START` (2026-08-01) are ignored** — Foxtrot took the contract over on 8/1 and the earlier rows are the previous vendor's service history, backfilled to seed compliance windows (they all carry zero revenue). JSX stations therefore read zero for May–July on purpose. **`SNA PRIV` and `SCF PRIV` do not exist in Paylocity yet**, so BUR-SNA counts only BUR's labor and SCF reads 0 worked hours until those codes are created. JSX services are all `aircraft: false`, so those stations show no "Avg aircraft/day" KPI — a debrief row can carry several services, so summing the flags would not be a plane count.
- **Manager grouping still matches on the first 3 letters of the sheet name**, so `BUR-SNA` inherits BUR's managers only (SNA's differ) and `TEB-HPN` inherits TEB's (same as HPN's, so no difference there).
- **Fixed budgets can be shaped by weekday.** A fixed service may carry a `dow_shape` in `station_overrides.json` (FLL FAC: Mon–Fri 49.3, Sat 2, Sun 0 — they work ~2 h Saturday and none Sunday). The workbook rate stays authoritative: a week is worth `rate x 7` and the shape is rescaled to that total, so editing the rate still moves the budget. Weekly totals and weekly averages are unchanged versus a flat rate; **month totals shift slightly**, because a month isn't a whole number of weeks and the trailing partial week may be all weekdays.
- **Ultra Debriefs holds two job types in one table** (`Service` = "Ultra Cleaning" or "Shroud Cleaning"). The workbook's Ultra COUNTIFS filtered on Location only, so every shroud job was also counted as an Ultra — inflating STL AA CAB by 9 jobs across the window (Aug 2026 fix). All Ultra/Shroud specs now filter `Service` explicitly, in `station_overrides.json` so an `extract_config.py` regen can't revert it. Only STL has shroud rows today; the CMH and DFW entries are latent protection.
- `catalog_extras.json` holds hand-added service templates (AA Turn/RON/Security, Ultra, APU Wash) merged over the generated catalog — regenerating `service_catalog.json` won't lose them. `station_overrides.json` holds per-station exceptions: STL AA CAB / PVD AA CAB labor dist codes match their sheet names (not the `<CODE> CABIN` convention), and STL AA CAB's "Shroud Cleaning" counts from AA_Debriefs, not the Ultra workbook DFW uses.
- The build intentionally **fixes four workbook bugs**: CMH counting LIT's Envoy aircraft (days 2–31), IAH budget rates shifted one row, BDL week-4 hardcoded rates, FLL/TUS FAC days 29–31 pulling MLB FAC's worked hours. Expect those cells to differ from Excel until the workbook is fixed.
- Data window starts 2026-05-01 (`WINDOW_START`), mirroring the hardcoded cutoff in the workbook's M queries.
- **Master page** ("All Locations — Overview", the default landing view): company-wide MTD KPIs, a per-station table, and By Regional Manager groups. Manager mapping comes from `Power BI Data Sources/Location Management.csv`, matched **blindly on the first 3 letters** of the location on both sides (the airport code) — per the owner's instruction. A location with multiple managers in the CSV appears under each of them, so manager totals overlap and do not sum to the company total.

## Secrets

Same three as `envoy-compliance-tracker` (see that repo's CLAUDE.md): `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET` (Foxtrot Report Automation app — secret expires every 24 months; renew in Entra and update here too).

## Local development

`python build_pulse.py` with no env vars reads the synced Data Hub folder directly. `PULSE_DATA_DIR=sources python build_pulse.py` after `python fetch_sources.py` mimics CI. Open `index.html` — no server needed.

## Org migration (2026-08-19)

Canonical repo: **Foxtrot-Aviation-Services/foxtrot-labor-pulse-04f16509** (private; the Pages
site is public at `foxtrot-aviation-services.github.io/foxtrot-labor-pulse-04f16509/`). The old
`sam-kosco.github.io/foxtrot-labor-pulse-04f16509/` URL stays live via a same-named mirror repo
on Sam's personal account, force-synced by this repo's "Mirror to legacy
URL" workflow (deploy key in `MIRROR_DEPLOY_KEY`). The mirror has Actions
DISABLED — never push to it or run anything there. Retire the legacy URL
(delete the mirror repo + mirror.yml) once the Foxtrot Platform rollout
replaces old links.
