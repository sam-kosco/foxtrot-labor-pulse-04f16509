# One pulse, one workbook — plan (drafted 2026-10-08, nothing built)

Sam, 2026-10-08: *"break down the artificial barrier between pulse sheets in
different verticals. One unified workbook, one locations page, and just have
the individual location pages have the different formats that currently
exist. Combine the different permissions, let everybody see all the pulse
pages for now."*

The split today is an accident of history, not a fact about the work: the
commercial pulse grew out of an Excel sheet and the MRO board was built
native, so a location's vertical decides which **workbook** configures it,
which **repo** renders it, and which **permission** opens it. None of that is
anything an ops person should have to know. This collapses all three, while
keeping the two page FORMATS, which are genuinely different — a commercial
station is budget-vs-worked per day, an MRO location is a drag-and-drop
forecast calendar.

**41 locations today**: 33 in `Service Budgets.xlsx`, 8 in
`Private and MRO Pulse Locations.xlsx`. No sheet-name collisions.

---

## 1. One workbook

Merge into **`Pulse Sheets/Pulse Locations.xlsx`**, one sheet per location,
41 sheets. Both workbooks are already "one sheet per location", so this is a
reshape, not a redesign.

### The uniform sheet shape

Every sheet gets the same two-part layout the MRO sheets already use:

```
A/B  key/value config, read generically, terminated by a blank key
C+   the vertical's own table
```

| | A/B keys | C+ table |
|---|---|---|
| Commercial / Facility | `Vertical`, `Labor Distribution`, optional knobs | `Service` \| `Budget` |
| Private / MRO | `Vertical`, `Labor Distribution`, `Hourly Goal`, `Facility Hours`, `Premium Hourly Goal`, `Premium Services`, `Layout`, `Feed`, `Attribution` | `Service` × plane-type price matrix |

`Vertical` is what a reader branches on, and it already exists on the MRO
side (`engine/mro.py` defaults it to `mro`). The MRO parser reads its A/B
block **generically** — "a new knob needs no code change here" — so the
commercial sheets slot into the same front-end for free.

### What this buys beyond tidiness

- **Labor dists become declared, in one place.** Today a commercial
  station's dists are implicit: a convention (`<code> CABIN` / `<code> FAC`)
  with `labor_keys` overrides hidden in `station_overrides.json`. Putting
  `Labor Distribution` on every sheet makes the thing that actually defines
  a location visible to the person editing it.
- **`assert_no_dist_overlap` can finally cover everything.** It lives in
  `build_mro_hours.py` and hard-fails when an MRO sheet claims a dist the
  commercial config also counts — but it only properly sees one side's
  dists. One workbook makes double-counting a lookup, not an audit.
- One file to fetch, one to permission, one to open when a location changes.

### Migration

33 commercial sheets get an A/B header inserted and their Service/Budget
table moved to C/D; the 8 MRO sheets copy across unchanged. Do it as **zip
surgery** (rewrite only the sheet XML plus `workbook.xml`, copy every other
part byte-for-byte) — a plain openpyxl round-trip silently drops all nine of
SharePoint's `customXml` parts, verified 2026-10-07. The workbook has no
tables, macros, drawings or ActiveX, so nothing else is fragile.

### Decisions to make before building

- **Name.** `Pulse Locations.xlsx` says what it is now that it is more than
  budgets — but `Service Budgets.xlsx` is wired into the Power Automate
  refresh buttons and whatever flow points at it. Renaming means re-pointing
  those once. **Recommend: new name.**
- **Three sheets share an airport with a different vertical**: `BNA AA` and
  `BNA FAC` vs `BNA`, and `TUS FAC` vs `TUS MHI`. In one list that only
  reads clearly if the MRO sheets are named for what they are — suggest
  `BNA MRO`, keep `TUS MHI`. Renaming an MRO sheet renames its schedule
  file (`MRO Schedules/<name>.json`), so it is a migration, not a cell edit.

---

## 2. One locations page

`#labor` becomes a native list of all 41 locations, no vertical filter and no
tabs. Clicking one opens the format that location actually has:

```
#labor                  Locations — every location, both verticals
#labor/pulse/<station>  the commercial pulse, deep-linked in the iframe
#labor/mro/<location>   the native Private/MRO calendar
```

- The list, its topline and a per-row vertical chip come from ONE read of
  the unified workbook — no stitching two sources, which is what made the
  previous attempt's topline awkward.
- **Budgets are comparable across verticals** (established 2026-10-07):
  `mro.month_to_date_budget()` prices each scheduled job at its revenue rate
  and spreads it over working days, month-to-date on the same elapsed-days
  rule the commercial side uses (day 1 through *yesterday*). So one
  Budgeted / Worked / Variance column set covers all 41.
- The commercial pulse needs `?station=NAME` and `?nav=0` (open on that
  station, hide its own picker so the list owns navigation). **That patch is
  already written and verified** — saved at
  `scratchpad/killed-reorg/pulse-deeplink.patch`.
- Keep `#labor/mro` working bare but unlinked, as before.

**Order matters**: pulse-web ships the deep link FIRST, platform second. In
between, a location link lands on the remembered station with the pulse's own
picker visible — degraded, not broken.

---

## 3. One permission

Retire `MRO Labor`; `Labor` covers both verticals and is granted to everyone.

- `MODULES` loses `MRO Labor`; `for_user` loses its special case (today
  `Labor Admin` doubles as MRO's admin tier).
- **Matrix migration, one-shot**: grant `Labor` to every existing row at the
  moment the `MRO Labor` column disappears from disk. Keyed to that
  disappearance, so it fires exactly once and a later revoke sticks — the
  `EVERYONE_MODULES` backfill cannot do this, because it only fires when a
  column first APPEARS and `Labor` has been on disk for months. (Built and
  tested against a Data Hub copy 2026-10-05: `MRO Labor` dropped, Labor
  145 → 150, `Labor Admin` preserved, Quality untouched, revoke verified
  sticky.)
- **Back up `Permissions.csv` before the first boot** — the migration drops
  the column and its 45 grants.
- `mro.auto_access()` becomes dead once Labor is universal; delete it rather
  than leave it uncalled.

### The one open question: seeing vs editing

"Let everybody see all the pulse pages" is unambiguous for the commercial
pulse, which is read-only and already a public Pages site. The MRO board is
**read-write**, and `POST /api/mro/schedule/<loc>` has no per-location
scoping — the permission is the whole gate. Granting Labor to all 150
accounts therefore lets anyone edit any of the 8 forecast calendars, up from
51 people today.

**Recommend** shipping this together with the rule from 2026-10-07: view open
to everyone; **edit** requires your org-chart scope to intersect the
location's `Labor Distribution`, plus Labor Admins. That was built and
verified end to end (AS/OM/GM/RM/Director all filter upward; an AS edits
their own location and gets 403 on another). It is ~40 lines, and the
unified workbook makes it cleaner — the dists that define a location are
right there on the sheet.

If "for now" means open it up and scope later, that is worth deciding
knowingly rather than by omission.

---

## Decided 2026-10-10 — naming and the location set

**Rule: a location is named after the labor distribution it counts.** Four
pages keep compound names because they genuinely pool dists, and two split
one dist between them:

| keep as-is | why |
|---|---|
| `BUR-SNA`, `DFW-DAL`, `FLL`, `TEB-HPN` | pool two dists each |
| `CLE DAY`, `CLE NIGHT` | split one dist (`CLE AA`) by shift window |

**`IAH` splits into `IAH CABIN` and `IAH FAC`** — the only page pooling a
cabin and a facility crew at the same airport. Note this also fixes a
standing trap: IAH is a commercial sheet today, so its FACILITY crew's hours
take the 12-hour shift-back. Split, `IAH FAC` gets plain calendar-day
attribution for free.

**`STL AA ULTRA` splits out of `STL AA CAB`**, carrying the **Ultra Cleaning
and Shroud Cleaning** services specifically. Those 6 people's hours are
counted nowhere today while their budget sits on STL AA CAB, which makes
STL AA CAB read over budget.

**Renames** (one-to-one dist, name differs — 28 of them). The bulk are the
`<CODE>` → `<CODE> CABIN` convention (BDL, CAK, CLT, CMH, CVG, DAY, DCA,
GSP, JFK, LIT, MCO, ORF, SGF, XNA) and the MRO sheets (AFW → `AFW EMB`,
BNA → `BNA MRO`, JAX → `JAX MRO`, MCN → `MCN EMB`, PVU → `PVU PRIV`,
RFD → `RFD MRO`, SLN 1V → `SLN MRO`, TUS MHI → `TUS MRO`). One-offs worth
noticing: **`BNA AA` actually counts `BNA CABIN`**, `STL AA FAC` counts
`STL FAC`, `TYS` counts `TYS PSA`, `SCF` counts `SCF PRIV`, and **`DFW WIDE`
counts `DFW WIDEBODY`** — the rule undoes the display rename from 2026-09-28,
so the `display` key in `station_overrides.json` can go with it.

Renaming an MRO sheet also renames its schedule file
(`MRO Schedules/<name>.json`) — a migration, not a cell edit.

### Locations still to create

`BNA PRIV`, `CAK PRIV`, `CLE PRIV`, `CVG FEAM`, `DTW PRIV`, `ILN MRO`,
`MLB STS`, `TPA PRIV` — real operations with no page today. **`LCQ FAC` is
built and live** (2026-10-10). **`MKE MRO` is dead and counts for nothing.**
`CAK HQ` is head office and gets no page.

> Sam's 2026-10-10 list named `MKE MRO` among the real locations one sentence
> after calling MKE dead — read as the latter. (`CGF FEAM` in that list is
> `CVG FEAM`.) Worth confirming once.

### The facility goal rate

**$30 of revenue per worked hour**, used to turn a fixed-revenue contract
into a daily budget: `monthly revenue / 30 calendar days / 30`. It is not
recorded as a constant anywhere — derived from the only two facility
locations with fixed monthly revenue in `ERP/Fixed Monthly Revenue.xlsx`
(FLL implies $29.98/h, MLB $31.09/h). **Writing it down somewhere canonical
is part of this work** — probably a `Goal Rate` key on the sheet, which the
unified A/B header makes natural, and which would let the other
fixed-revenue facilities be budgeted the same way instead of by hand.

## Corrected 2026-10-11 — MRO locations are not a different kind of thing

An earlier draft said an MRO location "has no budget-vs-worked day grid to
render". **That was wrong.** They have exactly the same two series the
commercial pulse has, per day:

- **budgeted** — `mro.daily_budgets()`: each scheduled job priced at its
  revenue rate (premium services on the premium goal), spread evenly over
  the job's working days, plus the flat facility allowance.
- **worked** — `mro_hours.json`, the same `worked_hours()` contract the
  commercial build uses, published by the same refresh.

Sampled 2026-10-01..10 to be sure: TUS MHI runs 54.6–194.8 h/day budgeted
against 0–96.3 worked; PVU 15.2–102.7 against 4.5–79.9.

So the unified list is **apples to apples at day resolution**, not just a
month-to-date approximation — every one of the 41 locations has budgeted,
worked and variance on the same terms. There is no second-class row.

**What differs is the destination, not the data.** Clicking a commercial or
facility location opens the pulse sheet; clicking a Private/MRO location
opens the drag-and-drop scheduler, because that is how those locations are
actually operated — the crews forecast there, and it is where their budget
comes from in the first place (Sam, 2026-10-11).

A consequence worth noting: an MRO location *could* be rendered as a pulse
sheet, and the data would be correct. That is a presentation choice left
open, not a constraint. The list does not have to care.

## Decided 2026-10-10 — one endpoint per location

Sam: *"Each location page should be its own endpoint so I can scope access
easily in the future."* So the route and the API are keyed by location, not
by vertical:

```
#labor                      the "Pulse Locations" list
#labor/<location>           that location's page, whichever format it has
#labor/mro[/<location>]     the Private/MRO board — KEPT as its own tab
GET /api/labor/locations    the list
GET /api/labor/location/<name>   one location's payload — the scoping seam
```

**Two tabs, not one** (Sam, 2026-10-11). The main tab is **Pulse Locations**
— every location, both verticals. The **Private/MRO Labor** tab stays beside
it, so the board has a front door of its own and the `#labor/mro` links
people already use keep working. The earlier draft retired that tab; it
shouldn't be. Two routes to the editable board is the point, not an
accident.

One endpoint per location means a future per-location gate is a decorator on
one function, not a redesign. The route stops encoding the vertical
(`#labor/pulse/...` / `#labor/mro/...` both go), which is the point — a
location's format is a property of the location, not of its URL.

### Reaching the scheduler from a pulse page — BUILT 2026-10-11

A Private/MRO location renders as a read-only pulse page, so it needed a way
out or it was a dead end: its numbers come from a forecast you cannot reach.
Every MRO/Private page now carries an **"✎ Edit the forecast"** link to
`#labor/mro/<location>` on the platform, with `target="_top"` so it escapes
the iframe when the platform embeds the page. Commercial pages don't show it.

This matters most for the standalone pulse URL, which is how most of ops
reaches the page today and which has no platform chrome around it at all.

### The catch worth knowing before you rely on it

**The commercial pulse is a public page.** `pulse_data.json` on the Pages
site answers **HTTP 200 with no authentication** — every location, every
day, to anyone with the URL (confirmed 2026-10-10; CLAUDE.md says as much:
"treat the URL as the access control"). The platform embeds it in an iframe.

So per-location endpoints make the *platform* scopeable while the data
behind the commercial half stays world-readable. The endpoints are still
worth building now — they are the right shape and they cost little — but
**scoping will not actually scope anything until the commercial pulse stops
being served from a public Pages site.** Two honest ways out when that day
comes:

1. **Render commercial locations natively**, with the platform fetching
   `pulse_data.json` server-side and the endpoint returning only the
   requested location's slice. This is the stated end-state anyway —
   `static/platform.js`'s own header says the embedded sections "will be
   re-rendered natively later". Cost: porting the week grid, the group
   rows and the tooltips into the platform.
2. **Accept that scoping is advisory** for commercial locations and real
   only for Private/MRO (already native and server-rendered).

Nothing in this release forecloses either. Building the endpoints now is
what makes option 1 a later afternoon rather than a rewrite.

## Order of work

1. Build the unified workbook and migrate (zip surgery), leaving both old
   files in place untouched.
2. Point `build_pulse.py`, `build_mro_hours.py` and `engine/mro.py` at it,
   behind a fallback to the old paths so nothing breaks mid-flight.
3. Ship the pulse-web deep link.
4. Platform: `/api/labor/locations` over the one workbook, the Locations
   list, the routing, one tab.
5. Permissions collapse and migration.
6. Retire the two old workbooks and `make_budget_workbook.py`.

Steps 1–2 are invisible to users and independently verifiable: the pulse must
build byte-identical output from the new workbook before anything else moves.

## Salvage from the killed plan

`scratchpad/killed-reorg/` holds the previous attempt — the Locations list
(`static/labor.js`), the routing, the permissions collapse, `_mro_can_edit`,
`mro.month_to_date_budget`, the pulse deep link, and the
**Labor Dist → Economic Unit → Region mapping** with its validation (two
dists missing from it, five with nobody booked, the blank-unit trap, AFW
7 → 55 editors). That directory is session-scoped — anything worth keeping
should be lifted out before it disappears.
