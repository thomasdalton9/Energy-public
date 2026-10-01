# Energy data repo - working conventions

Standing instructions from the repo owner. Follow these on every change.

## Branches
- Develop on `github-access-energy`; production is `main`. Every change goes to BOTH branches.
- External data sites are mostly blocked from the Claude sandbox: run pulls in GitHub Actions.

## Spreadsheets and charts
- **Every xlsx we pull must include native Excel charts of its data**, redrawn by the
  script on every run (openpyxl charts inside the workbook, not just a separate PNG).
  - Storage / reservoir / water-level series: AGSI-style water-year chart via
    `water_year_chart.add_water_year_chart()` (5-year min-max band, 5Y average,
    previous and current water year).
  - Other time series: line or stacked chart of the main columns on a chart sheet.
- Storage/level/seasonal charts use an **Oct-Sep water year**, not Jan-Dec.
- Every new or updated PNG chart is committed to `output/` on main and sent to the owner.
- Units/notes tab via `xlsx_notes.write_workbook()` (atomic write).

## Pull scripts
- Incremental: backfill gaps only, don't re-pull complete history each run.
- South America gas demand by sector: data from 2021 only.
- Ireland gas: GNI transparency pages from 2026-03-31 on (open data covers earlier);
  ENTSOG is for validation only.

## Housekeeping
- One-off discovery/inspect/probe/test scripts go in `discovery_archive/`
  (with their manual-only workflows in `discovery_archive/workflows/`).
