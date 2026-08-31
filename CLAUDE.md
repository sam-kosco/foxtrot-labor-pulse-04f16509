# foxtrot-labor-pulse

**Owner:** Samuel Kosco — Data Analyst, Foxtrot Aviation Services
**Purpose:** Read-only Labor Pulse dashboard for ops — the web replacement for opening `Live Commercial Pulse Sheet.xlsx`. Budgeted vs worked labor hours per station per day, with weekly averages and variance.

The page is intentionally uneditable: ops get a link, not the workbook. The repo name carries a random suffix and the page carries `noindex` because the Pages site is public — treat the URL as the access control and share it only internally. Since the org migration the CODE is private (org repo); only the rendered site is public. The platform embeds this site and its JSON feeds the platform homepage stats.

## How it works

1. `.github/workflows/refresh.yml` is poll-triggered (2026-08-21; dispatcher-fired since 2026-08-31 — GitHub's cron is fully retired org-wide): the Foxtrot Platform's dispatcher fires a `workflow_dispatch` every 30 min, 6:00 AM–6:30 PM Eastern (manifest `Monitoring/schedules.json`). Each fire is a light poll — `check_sources.py` compares each source's Graph lastModifiedDateTime against `sources_meta.json` (committed at every build) and a full rebuild runs only when something changed — so the main refresh follows the ~6 AM Eastern Paylocity relay within minutes and same-day debrief edits surface within the poll interval. Tick the `force` input (the Actions-tab default for humans) to rebuild unconditionally. The workflow name ("Daily Pulse Refresh") is load-bearing — mirror.yml's workflow_run matches on it.
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
- **JSX stations** (`BUR-SNA`, `TEB-HPN`, `SCF`, and DAL's jobs inside `DFW-DAL`) read `Power Flows/Debriefs/JSX Debriefs.xlsx`, table `Debriefs` on Sheet1 — one row per tail with 0/1 columns for RON, Interior Detail, Exterior Detail and Carpet Extraction, keyed by `Service Location`. **Exterior details are priced by airframe** (Aug 2026): the build splits the single `Exterior Detail` flag into `Exterior Detail ATR` (Plane Type contains "ATR") and `Exterior Detail 145` (everything else, deliberately including the odd unreadable `#N/A` Plane Type - the build logs those rather than dropping them). Biohazard was retired as a service; the column still exists in the debrief but nothing reads it. Two sheets each cover two airports (Excel forbids `/` in sheet names, hence the hyphen); their specs list one COUNTIFS per airport and the build sums them. **Rows dated before `JSX_START` (2026-08-01) are ignored** — Foxtrot took the contract over on 8/1 and the earlier rows are the previous vendor's service history, backfilled to seed compliance windows (they all carry zero revenue). JSX stations therefore read zero for May–July on purpose. **`SNA PRIV` and `SCF PRIV` do not exist in Paylocity yet**, so BUR-SNA counts only BUR's labor and SCF reads 0 worked hours until those codes are created. JSX services are all `aircraft: false`, so those stations show no "Avg aircraft/day" KPI — a debrief row can carry several services, so summing the flags would not be a plane count.
- **`DFW-DAL` is one page for two airports** (Aug 2026): the DFW team does much of the DAL work, so budgets and labor are pooled — labor dists `DFW CABIN` + `DAL PRIV`, and DAL's JSX services sit on the DFW-DAL sheet. `{LOC}` resolves to DFW (first 3 letters), which is right for the Envoy/Frontier/Regional services; the Ultra and JSX services are pinned explicitly in `station_overrides.json` because Ultras must filter `Service` and the JSX jobs are logged under DAL. The old `DFW` entry was removed from stations.json so the build stops logging "dropped from the pulse"; nothing referenced it, since DFW-DAL resolves its Envoy/Frontier/Regional services from `service_catalog.json` templates. Re-running `extract_config.py` against the legacy Excel pulse workbook would re-add it — drop it again if so.
- **Pre-launch labor can be gated per labor dist, not just per station.** A pooled station uses `labor_from_first_debrief: {"<LABOR DIST>": [<service keys>]}` — the named cost centre counts nothing before the first debrief of those services, while the station's other cost centres are untouched. DFW-DAL gates `DAL PRIV` on the JSX services (first debrief 2026-08-19). Using the station-wide `hours_from_first_debrief` here would have erased ~11,900 h of real DFW history; without any gate, ~377 h of DAL's implementation/imputed-salary time would have scored against DFW. With the gate the merge is exactly additive.
- **`CLE DAY` / `CLE NIGHT` split one labor dist by shift** (Aug 2026). Both draw on the single `CLE AA` cost centre; `shift_window` in `station_overrides.json` assigns each punch by its local punch-in hour — DAY takes 07:00-18:59, NIGHT takes 19:00-06:59. NIGHT additionally carries **all** CLE AA salaried imputation (DAY has an explicit empty `salary_keys`, which the build now distinguishes from an absent one). Day attribution follows the window: a wrapping (overnight) window keeps the punch-in-minus-12h rule that groups a night shift onto the day it started, while a daytime window uses the plain calendar day — otherwise every 7 AM start would be credited to the previous day. Consequence: totals are conserved but a few hours move between days versus the old combined CLE AA, and May gains ~4 h that previously shifted back past `WINDOW_START`. Ultra, Shroud and Frontier work all belongs to NIGHT; the facility budget is split on the sheets (3.75 + 8 = the former 11.75). Punches with no punch-in time would fall outside both windows — there are none today, and the holiday rows that lack one are already excluded upstream.
- **Pre-launch labor is suppressed.** Stations flagged `hours_from_first_debrief` in `station_overrides.json` (the four JSX sheets and BNA AA) count no labor — hourly, open-shift estimates, or salaried imputation — before the date of their own first debrief. Those hours are implementation and training for a contract that hadn't started, and mostly showed up as pure salaried imputation (5.71 h/day) against a station doing no billable work. The cutoff is **derived from the debrief data each build** (`first_debrief_dates`), not hardcoded, so it moves automatically as debriefs land or are backfilled — and the build logs each station's date. Note this interacts with `JSX_START`: JSX debriefs before 8/1 are dropped first, so a JSX station's cutoff is necessarily its first August debrief.
- **Manager grouping still matches on the first 3 letters of the sheet name**, so `BUR-SNA` inherits BUR's managers only (SNA's differ) and `TEB-HPN` inherits TEB's (same as HPN's, so no difference there).
- **Fixed budgets can run on set weekdays.** A fixed service may carry `dow_days` (the weekdays the workbook rate applies to; all other days 0) and optionally `dow_overrides` (per-day hours replacing the rate on named days) in `station_overrides.json`. Both weekday-based items use this one mechanism: **FLL FAC** runs 49.3 h Mon-Fri with a 2 h Saturday half-shift and no Sunday (weekly total 248.5), and **STL AA CAB's Mail Running** is 8 h Mon-Fri with nothing at weekends. The workbook rate means "hours on a normal working day", so editing Service Budgets.xlsx still moves the budget. Note month totals for these services depend on how many working days the month holds, so they will not equal rate x days.
- **Ultra Debriefs holds two job types in one table** (`Service` = "Ultra Cleaning" or "Shroud Cleaning"). The workbook's Ultra COUNTIFS filtered on Location only, so every shroud job was also counted as an Ultra — inflating STL AA CAB by 9 jobs across the window (Aug 2026 fix). All Ultra/Shroud specs now filter `Service` explicitly, in `station_overrides.json` so an `extract_config.py` regen can't revert it. Only STL has shroud rows today; the CMH and DFW entries are latent protection.
- `catalog_extras.json` holds hand-added service templates (AA Turn/RON/Security, Ultra, APU Wash) merged over the generated catalog — regenerating `service_catalog.json` won't lose them. `station_overrides.json` holds per-station exceptions: STL AA CAB / PVD AA CAB labor dist codes match their sheet names (not the `<CODE> CABIN` convention), and STL AA CAB's "Shroud Cleaning" counts from AA_Debriefs, not the Ultra workbook DFW uses.
- The build intentionally **fixes four workbook bugs**: CMH counting LIT's Envoy aircraft (days 2–31), IAH budget rates shifted one row, BDL week-4 hardcoded rates, FLL/TUS FAC days 29–31 pulling MLB FAC's worked hours. Expect those cells to differ from Excel until the workbook is fixed.
- Data window starts 2026-05-01 (`WINDOW_START`), mirroring the hardcoded cutoff in the workbook's M queries.
- **Master page** ("All Locations — Overview", the default landing view): company-wide MTD KPIs, a per-station table, and By Regional Manager groups. Manager mapping comes from `Power BI Data Sources/Location Management.csv`, matched **blindly on the first 3 letters** of the location on both sides (the airport code) — per the owner's instruction. A location with multiple managers in the CSV appears under each of them, so manager totals overlap and do not sum to the company total.

## Private/MRO hours feed (mro_hours.json)

`build_mro_hours.py` (runs in the same refresh, after `build_pulse.py`) publishes
`mro_hours.json` beside `pulse_data.json`: per-day worked hours for each Private/MRO
pulse location, consumed by the Foxtrot Platform's Private/MRO Labor calendar (see
`MRO_PULSE_PLAN.md` in the foxtrot-platform repo).

- Locations come from `Pulse Sheets/Private and MRO Pulse Locations.xlsx` — one sheet
  per location; columns A/B are a key/value block read until the first blank key. This
  build uses `Labor Distribution` (comma-split for multi-dist) and the optional
  `Attribution` knob; the price matrix from column E is the platform's concern.
- Worked hours call the **same `worked_hours()` the commercial build uses**
  (extracted from `build_month` 2026-08-31; the calculation contract — paid-duration
  punches, holiday/PTO exclusion, dist repair, open-shift estimation, 40/7 salaried
  imputation — lives in that one function). One deliberate difference: **day
  attribution defaults to the plain calendar day** (owner decision 2026-08-31 — MRO
  crews work day shifts; no 12 h shift-back). A sheet can opt back in with an
  `Attribution` row valued `shift`. No launch gating — MRO locations have no debriefs.
- **A labor dist may feed exactly one pulse**: the build hard-fails if an MRO sheet
  claims a dist the commercial config already counts (`BUR PRIV`/`DAL PRIV` are the
  live hazard — both belong to commercial stations today). In the workflow this build
  is `continue-on-error` so an MRO failure can't stale the commercial page; a final
  step still turns the run red.
- Window: last 70 days through today.

## Adding a source

`sources.py` is the **single list** of Data Hub files, shared by `fetch_sources.py`
(downloads them in CI) and `build_pulse.py` (`check_sources_present` fails fast by
name if one is absent). Wire any new source in there — adding a reader to the build
without adding the path here is what broke every scheduled refresh on 2026-08-21.

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
