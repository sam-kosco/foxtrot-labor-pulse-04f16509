"""Canonical list of Data Hub files the pulse reads.

Single source of truth shared by fetch_sources.py (which downloads them in CI)
and build_pulse.py (which validates they are all present before building).
Keeping one list means a new source can't be wired into the build and then
silently missed by the fetch — which is exactly how the 2026-08-21 refresh
broke: JSX Debriefs.xlsx was added to the build but not to the download list,
so every scheduled run failed with FileNotFoundError.

Paths are relative to the Data Hub drive root.
"""

SOURCE_FILES = [
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
    "Power Flows/Debriefs/JSX Debriefs.xlsx",
    "Power BI Data Sources/Location Management.csv",
    "Pulse Sheets/Service Budgets.xlsx",
]

# Basenames as they land in the flat CI sources/ directory.
SOURCE_NAMES = [p.rsplit("/", 1)[-1] for p in SOURCE_FILES]
