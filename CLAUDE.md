# Energy data repo - working conventions

Standing instructions from the repo owner. Follow these on every change.

## Branches
- Develop on `github-access-energy`; production is `main`. Every change goes to BOTH branches.
- External data sites are mostly blocked from the Claude sandbox: run pulls in GitHub Actions.
- Pulls run in the PUBLIC repo thomasdalton9/Energy-public (free Actions minutes); make code changes there.
  Its main is copied into the private thomasdalton9/Energy (main + github-access-energy) on the 1st and 15th (23:50 UTC),
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
- `europe/EUROPE_MASTER.py` -> `europe_master.xlsx`: Europe - the ENTSO-E countries (EU/EEA, UK excluded, plus the Balkans) plus
  Ireland EirGrid, Turkey, Cyprus and the Rhine at Kaub. Gas Dashboard (GIE AGSI+ storage and ALSI LNG, Ireland GNI),
  "Dashboard - Power" (generation, capacity, day-ahead prices, net imports, and a supply/demand balance per country: generation +
  net imports + pumped storage/batteries vs load) and "Dashboard - Capacity factors". Country list and bidding zones live in
  `europe/europe_countries.py` (add a country there and the pulls, registry and master pick it up). ENTSO-E is the primary
  power source for Europe (it is the TSOs' own statutory reporting, and matched SMARD/RTE within 0.5% in a settled-week check),
  labelled as such; replace it per country with a national feed where that is better (e.g. CGES for Montenegro, Elexon for GB,
  neither yet pulled). Gas pipeline imports/exports (ENTSOG) are still to add.
- Every dashboard chart shows its source. Prefer raw sources (grid operators, ministries, statistics offices);
  Ember is a fallback only for countries with no raw feed, labelled as such.

## Pull scripts
- Schedules: pulls whose source keeps history run on the 1st and 15th of each month (they backfill every missed
  day), to save Actions minutes. Only sources with no history run daily or more often: the Argentina AIC snapshot
  (argentina_aic_snapshot.yml; the full Argentina hydro pull is 1st/15th), Canada IESO (rolling 'today' XML),
  Ecuador CENACE daily, LNG feedgas (TC keeps no history), Turkey EPIAS ('today' only), Australia gas hub prices
  (au_sttm_prices.yml: AEMO STTM report INT651 holds about a week, DWGM about 14 days). Masters rebuild on the
  1st/15th evening, then the private-repo sync (23:50). A new pull follows the 1st/15th schedule unless its
  source has no history.
- Incremental: backfill gaps only, don't re-pull complete history each run.
  The committed workbook IS the history store: each run reads it and fetches only periods not saved yet (plus a
  short revision window). Sources that only publish whole files (e.g. MBIE webtables, Hydro Tasmania) are
  downloaded only when a new release/Last-Modified appears (recorded on the Units sheet).
- South America gas demand by sector: data from 2021 only.
- Ireland gas: GNI transparency pages from 2026-03-31 on (open data covers earlier);
  ENTSOG is for validation only.

## Housekeeping
- One-off discovery/inspect/probe/test scripts go in `discovery_archive/`
  (with their manual-only workflows in `discovery_archive/workflows/`).
