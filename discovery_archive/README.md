# Discovery archive

One-off scripts used to find and check data sources before the real pull
scripts were written: `*_DISCOVERY`, `*_INSPECT`, `*_PROBE`, `*_TEST`, `*_PRINT`,
`*_DEPTH_CHECK` and similar. Each one's docstring records what it found;
the production scripts cite them ("Found via X_DISCOVERY.py").

- Scripts keep their original folder layout under this directory
  (e.g. `south_america/ARGENTINA_GAS_DISCOVERY.py`).
- `workflows/` holds their manual-only GitHub Actions workflows. They are
  out of `.github/workflows/`, so they no longer appear in the Actions tab.
  To re-run one, copy it back to `.github/workflows/` and fix the script
  path to point here (scripts that import `xlsx_notes` from the repo root
  also need one more `os.path.dirname` in their `sys.path` line).

Nothing scheduled depends on anything in this folder.
