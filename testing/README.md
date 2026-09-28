# Test zone

Runs any script from any branch on GitHub, the same way it runs on your
desktop, without touching the automated jobs (`americas/feedgas_ebb/`,
`south_africa/`).

## Run a script

Actions -> **Test zone** -> Run workflow:

- **code_branch** - the branch with the code, e.g. `claude/new-project-setup-3eipfq`
- **script** - path from the repo root, e.g. `europe/gie_storage.py`
- **args** - anything the script takes on its command line
- **packages** - extra pip packages (pandas, openpyxl, requests, xlrd,
  matplotlib are always there)
- **browser** - tick if the script uses Playwright

Files the script creates or changes are copied to
`testing/outputs/<script name>/` and committed; the run's log shows its
printed output.

## API keys - no code changes needed

Your scripts `import api_keys` (the gitignored `api_keys.py` on your PC).
On GitHub the workflow writes that file fresh for every run from the
repository's Secrets, then it is thrown away with the runner - it is never
committed.

Add each key once under **Settings -> Secrets and variables -> Actions ->
New repository secret**, using the same name as in `api_keys.example.py`
(`EIA_API_KEY`, `ENTSOE_API_KEY`, `GIE_API_KEY`, `EMBER_API_KEY`, ...). Keys
without a secret are left blank, exactly as in a fresh `api_keys.py`.
Every secret is also available as an environment variable.

`testing/check_keys.py NAME1 NAME2 ...` (code_branch `main`) reports which
keys a run can see - names and lengths only, never values.
