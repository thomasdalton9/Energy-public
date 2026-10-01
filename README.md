# Energy data

Scheduled pulls of gas and power data, run in GitHub Actions. Workbooks (with native Excel charts) are in
`output/Data and Chart Outputs/`, chart PNGs in `output/PNG Charts/`, and copies of the producing scripts in
`output/Python Scripts/`. Working conventions are in [CLAUDE.md](CLAUDE.md).

## South & Central America coverage

![South & Central America data coverage](output/PNG%20Charts/south_central_america_coverage_map.png)

Green: gas demand by sector, gas production/supply and power generation by type. Blue: power by type or a gas
demand split only. Grey: no data. All of it feeds `south_and_central_america_master.xlsx`; the map is redrawn by
`south_america/COVERAGE_MAP.py`.
