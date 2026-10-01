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
  - Charts have **no borders** (no chart-area or plot-area outline).
  - Chart dates are formatted **mmm/yy** (e.g. Jan/26); annual series show the year.
  - `add_charts.py` holds the per-workbook chart registry; each scheduled workflow runs it after the pull.
- Storage/level/seasonal charts use an **Oct-Sep water year**, not Jan-Dec.
- Outputs live in `output/Data and Chart Outputs/` (xlsx + png); a copy of each producing script is kept in
  `output/Python Scripts/` (refreshed by the workflow on every run).
- Every new or updated PNG chart is committed to `output/Data and Chart Outputs/` on main and sent to the owner.
- Units/notes tab via `xlsx_notes.write_workbook()` (atomic write).

## Master workbooks
- `south_america/SOUTH_AMERICA_MASTER.py` -> `south_america_master.xlsx`: Dashboard front page with every South
  America chart, a data tab per chart, raw data tabs, Sources tab. New South American datasets: add them to its
  DATASETS list (and add_charts.py REGISTRY).

## Pull scripts
- Incremental: backfill gaps only, don't re-pull complete history each run.
- South America gas demand by sector: data from 2021 only.
- Ireland gas: GNI transparency pages from 2026-03-31 on (open data covers earlier);
  ENTSOG is for validation only.

## Housekeeping
- One-off discovery/inspect/probe/test scripts go in `discovery_archive/`
  (with their manual-only workflows in `discovery_archive/workflows/`).
