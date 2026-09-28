# Test zone

Scratch space for trying scripts on GitHub, separate from the automated
jobs (`americas/feedgas_ebb/`, `south_africa/`), which this never touches.

## Run a script

Actions -> **Test zone** -> Run workflow:

- **script** - path from the repo root, e.g. `testing/check_keys.py`
- **args** - anything the script takes on its command line
- **packages** - extra pip packages (pandas, openpyxl, requests are always there)
- **browser** - tick if the script uses Playwright

Anything the script writes under `testing/` is committed back; the run's log
shows its printed output.

## API keys

Never put a key in a file in the repo. Add each one once under
**Settings -> Secrets and variables -> Actions -> New repository secret**
(e.g. `ENTSOE_API_KEY`). Every secret reaches the script as an environment
variable of the same name, masked in the logs:

```python
import os
key = os.environ["ENTSOE_API_KEY"]
```

On your PC, set the same variable (or keep a `.env` file - it is gitignored)
and the script runs unchanged.

`testing/check_keys.py NAME1 NAME2 ...` reports which keys a run can see
(names only, never values).
