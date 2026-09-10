# TestCase Validator

**Version: v1.0** — actively evolving, see [Roadmap](#roadmap--future-versions).

`TestCase_validator.py` takes your QA test case suite (Excel)
and your app's source code (ZIP), then evaluates every test case two
different ways and cross-checks the results:

1. **Static check** — searches your source code for text that matches
   each test case (quoted strings, error messages, emails, Module/Role
   keywords). No running app needed for this half.
2. **Live check** — opens a real Chromium browser (via Playwright), logs
   in, and clicks through each test's actual steps against your running
   app, then checks what's really on the page.

Each check is wrong in a different, predictable way — the static check
misses features worded differently in code; the live check misses
features it can't find a selector for. Combining them turns their
disagreement into a signal instead of noise, and every verdict gets a
**Confidence** rating (High / Medium / Low) so you know how much to
trust it.

The output keeps your **exact original Excel format** — same columns,
`Status` overwritten with the result, nothing added or removed from that
sheet. A new `Automated Evidence Log` sheet holds the full reasoning.

---

## Requirements

```bash
pip install openpyxl playwright --break-system-packages
playwright install chromium
```

(`playwright` and `chromium` are only needed for the live half — see
`--static-only` below if you want to skip them.)

Tested on Python 3.8+.

---
## Verify your installation

Before running the validator for real, confirm both dependencies are
actually working — `pip install` succeeding doesn't guarantee Playwright's
browser binary downloaded correctly, and that's the most common setup
failure people hit.

Save this as `check_setup.py` in the same folder as the validator:

Then run:

\```bash
python check_setup.py
\```

**Expected output:**
\```
openpyxl ok
chromium ok
\```

If you see both lines, you're ready to run the validator. If it errors
instead:

- `ModuleNotFoundError: No module named 'openpyxl'` or `'playwright'` →
  re-run `pip install openpyxl playwright --break-system-packages`
- `Executable doesn't exist at .../chromium.../chrome...` → the Python
  package installed but the browser binary didn't — run
  `playwright install chromium` (or `python -m playwright install chromium`
  if `playwright` isn't recognized as a command)

## Excel format

Your test case sheet needs a `Test Cases` sheet with (at least) these
column headers in row 1:

```
Test Case ID | Module | Test Scenario | Test Case Title | Preconditions |
Test Steps | Test Data | Expected Result | Priority | Type |
Role / Persona | Status
```

- **Expected Result** — put anything mechanically checkable in quotes,
  e.g. `The page shows 'Invalid email or password.'`. Unquoted narrative
  results (e.g. "the user sees a friendly dashboard") can't be matched
  precisely and land on `REVIEW`.
- **Test Data** — for Login test cases, use `Email: x@y.com / Password:
  secret123` so the live half can learn credentials automatically for
  other test cases that need to be logged in as that role.
- A `Summary` sheet is optional but recommended — the script appends a
  results block to it if present.

---

## Usage

```bash
# Both checks (default, recommended)
python TestCase_validator.py \
  --excel your_Test_Case_Suite.xlsx \
  --zip   your_source.zip \
  --base-url http://localhost:3000

# Let the script start your app itself
python TeseCase_validator.py \
  --excel your_Test_Case_Suite.xlsx \
  --zip   your_source.zip \
  --base-url http://localhost:3000 \
  --start-cmd "npm install && npm start" \
  --cwd frontend

# Static check only -- no running app, no Playwright needed
python TeseCase_validator.py \
  --excel your_Test_Case_Suite.xlsx --zip your_source.zip --static-only

# Live check only -- skip the source code search
python TeseCase_validator.py \
  --excel your_Test_Case_Suite.xlsx --zip your_source.zip \
  --base-url http://localhost:3000 --live-only
```

**First run on a new app? Do this** to watch it work on a small batch
before committing to a full run:

```bash
python TeseCase_validator.py \
  --excel your_Test_Case_Suite.xlsx \
  --zip   your_source.zip \
  --base-url http://localhost:3000 \
  --max-tests 10 --filter-module Login --headed --debug
```

`--headed` shows the actual browser; `--debug` prints every selector it
tries and whether it matched.

---

## How the verdicts combine

| Static says | Live says | Final | Confidence |
|---|---|---|---|
| Found in code | Confirmed on screen | **PASS** | High |
| Found in code | Couldn't confirm | REVIEW | Medium |
| Found in code | Not seen on screen | REVIEW | Medium — possible real bug |
| Partial match | Confirmed on screen | **PASS** | Medium |
| Partial match | Partial/uncertain | REVIEW | Low |
| Partial match | Not seen on screen | **FAIL** | Medium |
| Not found in code | Confirmed on screen | **PASS** | Medium — likely dynamic content |
| Not found in code | Couldn't confirm | REVIEW | Medium |
| Not found in code | Not seen on screen | **FAIL** | High |

---

## Getting real accuracy: `--selector-map`

Heuristic guessing (`input[type=email]`, button text matching, etc.) only
gets you so far. For real accuracy on **your** app, inspect it once in
dev tools and give the script exact selectors:

```json
{
  "fields": {
    "email": "#loginEmail",
    "password": "#loginPassword"
  },
  "buttons": {
    "Sign In": "button[data-testid='login-submit']",
    "Approve": ".approvals-card button.approve-btn"
  }
}
```

```bash
python TeseCase_validator.py ... --selector-map my_app_selectors.json
```

Any field/button name in the map is used **before** any heuristic guess.

---

## Speed

Static checking is near-instant (plain text search); the live half is
what takes time, since it's driving a real browser through each test.

| Mode | Approx. speed | 446-row suite |
|---|---|---|
| `--static-only` | seconds total | seconds |
| `--live-only` / combined (default) | ~2–3 sec per test case | ~20–25 minutes |

Actual speed depends on `--per-test-timeout`, how many candidate
selectors each step has to try before matching, and your app's own
response time. A `--selector-map` speeds things up too, since a mapped
field/button matches on the first try instead of falling through the
heuristic candidate list.

---

## Accuracy

There's no single accuracy percentage — it depends on your app's markup
and how many `Expected Result` cells use quoted literal text the script
can mechanically check (narrative-only results always land on `REVIEW`
by design, rather than risk a wrong PASS/FAIL). What you can rely on is
the **Confidence** rating, which reflects exactly how the two checks
behaved on that row:

| Confidence | What it means | How much to trust it |
|---|---|---|
| **High** | Static and live checks independently agreed (both PASS or both FAIL) | Most reliable — spot-check occasionally |
| **Medium** | One check gave a clear answer, the other was weak/absent, or they landed on opposite strong answers | Worth a look, especially FAIL/PASS disagreements |
| **Low** | Neither check found strong evidence either way | Needs a human — treat as un-evaluated |

In testing against a sample app, giving the live check a
`--selector-map` for its login form took every login test case from
`REVIEW`/Medium confidence straight to `PASS`/High confidence — the
selector map is the single biggest thing you can do to raise real
accuracy on your specific app.

---

## Output

- **`Test Cases` sheet** — your original columns, `Status` updated and
  color-coded (green = PASS, yellow = REVIEW, red = FAIL)
- **`Automated Evidence Log` sheet** (new) — per-row reasoning: static
  evidence (file:line), live steps parsed/executed/skipped, notes,
  screenshot path, and Confidence rating
- **`Summary` sheet** — a results block appended at the bottom
- **`screenshots/` folder** — one PNG per test case (live/combined modes)

Nothing is deleted or reordered in your original sheets.

---

## All options

| Flag | Default | Description |
|---|---|---|
| `--excel` | *required* | Path to the test case `.xlsx` |
| `--zip` | *required* | Path to the source code `.zip` |
| `--base-url` | required unless `--static-only` | URL the app is/will be reachable at |
| `--start-cmd` | none | Shell command to start the app (omit if already running) |
| `--cwd` | ZIP root | Subfolder inside the ZIP to run `--start-cmd` from |
| `--output` | `<excel>_COMBINED_RESULTS.xlsx` | Output path |
| `--sheet` | `Test Cases` | Sheet name holding the rows |
| `--static-only` | off | Skip the live browser run entirely |
| `--live-only` | off | Skip the static code check entirely |
| `--startup-timeout` | `60` | Seconds to wait for the app to come up |
| `--per-test-timeout` | `20000` | Max ms for any single wait/action |
| `--settle-ms` | `400` | Pause after navigations/clicks for SPA state to catch up |
| `--selector-map` | none | JSON file of CSS selector overrides |
| `--debug` | off | Print every selector attempt |
| `--screenshots-dir` | `screenshots` | Where per-test screenshots are saved |
| `--headed` | off | Show the browser instead of running headless |
| `--keep-server` | off | Don't stop the app process when done |
| `--max-tests` | none | Only run the first N selected rows |
| `--filter-module` | none | Only run rows from one `Module` |

---

## What this doesn't do

- **Confirm business logic is correct** — it checks that expected
  text/behavior appears, not that the underlying logic is right in every
  edge case.
- **Replace a human reviewer** — treat `PASS` as "worth spot-checking,"
  not "certified correct." `REVIEW` and Low-confidence rows are exactly
  the ones that need a person to look.
- **Understand unusual UI** — icon-only buttons, canvas-based UI, or
  heavily custom components will produce more `REVIEW`/`FAIL` results
  until you supply a `--selector-map`.

---

## Roadmap / future versions

v1.0 is the initial release. Planned for future versions (edit this list
as your own priorities take shape):

- Parallel test execution to cut down full-suite runtime
- Auto-generating a starter `--selector-map` by scanning the app's DOM
- Smarter multi-step / multi-role parsing (e.g. steps that navigate
  between pages or switch role mid-test)
- Retry / self-healing selectors when the DOM changes slightly
- An HTML/PDF summary report in addition to the Excel output
- CI/CD integration (e.g. a GitHub Actions workflow that runs this on
  every PR)

---

