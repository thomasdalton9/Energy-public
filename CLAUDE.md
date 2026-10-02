# Energy data repo - working conventions

Standing instructions from the repo owner. Follow these on every change.

## Branches
- Develop on `github-access-energy`; production is `main`. Every change goes to BOTH branches.
- External data sites are mostly blocked from the Claude sandbox: run pulls in GitHub Actions.
- Pulls run in the PUBLIC repo thomasdalton9/Energy-public (free Actions minutes); make code changes there.
  Its main is copied into the private thomasdalton9/Energy (main + github-access-energy) weekly, Monday 23:30 UTC,
  by .github/workflows/sync_to_private.yml (needs the PRIVATE_REPO_TOKEN secret). Never copy private-repo
  content into the public repo.

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
- Outputs: xlsx in `output/Data and Chart Outputs/`; PNG charts in `output/PNG Charts/`; a copy of each producing
  script in `output/Python Scripts/` (refreshed by the workflow on every run).
- Every new or updated PNG chart is committed to `output/PNG Charts/` on main and sent to the owner.
- Units/notes tab via `xlsx_notes.write_workbook()` (atomic write).

## Master workbooks
- `south_america/SOUTH_AMERICA_MASTER.py` -> `south_and_central_america_master.xlsx`: South America plus Central
  America (Guatemala-Panama + Belize; NOT Mexico) and the Caribbean: Trinidad & Tobago, Puerto Rico, Jamaica,
  Dominican Republic. Gas Dashboard + "Dashboard - Power & Hydro", a data tab per
  chart, raw data tabs, Sources tab. New datasets: add them to its DATASETS / RAW_POWER_DATASETS / HYDRO_DATASETS
  lists (and add_charts.py REGISTRY), plus a SOURCES entry (publisher + link).
- `americas/NORTH_AMERICA_MASTER.py` -> `north_america_master.xlsx`: United States, Canada, Mexico, same layout
  (reuses SOUTH_AMERICA_MASTER's code). Gas Dashboard + "Dashboard - Power"; add datasets to its DATASETS /
  RAW_POWER_DATASETS / OTHER_POWER_DATASETS / CAPACITY_DATASETS lists plus a SOURCES entry.
  US power generation comes from EIA-930 (balancing-authority data, current to yesterday) - owner's decision; don't
  switch the US regions to direct ISO feeds.
- Every dashboard chart shows its source. Prefer raw sources (grid operators, ministries, statistics offices);
  Ember is a fallback only for countries with no raw feed, labelled as such.

## Pull scripts
- Schedules: pulls whose source keeps history run WEEKLY (Monday; they backfill every missed day), to save
  Actions minutes. Only sources with no history stay daily or more often: Argentina hydro (AIC snapshot),
  Canada IESO (rolling 'today' XML), Ecuador CENACE daily, LNG feedgas (TC keeps no history), Turkey EPIAS
  ('today' only). Masters rebuild Mon + Thu evening. A new pull is weekly unless its source has no history.
- Incremental: backfill gaps only, don't re-pull complete history each run.
- South America gas demand by sector: data from 2021 only.
- Ireland gas: GNI transparency pages from 2026-03-31 on (open data covers earlier);
  ENTSOG is for validation only.

## Housekeeping
- One-off discovery/inspect/probe/test scripts go in `discovery_archive/`
  (with their manual-only workflows in `discovery_archive/workflows/`).
