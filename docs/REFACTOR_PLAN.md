# PharmParser — Refactor Plan

**Status:** proposal · **Scope:** whole codebase (~770 LOC across 24 Python files) · **Author:** generated from a full read of every source file at `b80debc`

---

## 1. Where the project stands

PharmParser scrapes drug prices from `tabletka.by` for a set of pharmacies, compares them, and
writes a formatted `.xlsm` workbook with VBA sort/filter buttons. It is driven by a CustomTkinter
desktop GUI.

The code works, but it has accumulated the shape of a script that grew into an app:

| | |
|---|---|
| Test files | **0** |
| CI / lint / typecheck config | **none** |
| Installable on Linux/macOS | **no** — `uv sync` fails outright |
| Pure, side-effect-free business logic | **~none** — every rule is inside a widget, a worksheet, or a COM call |
| Layering | `core` → `utils` → `customtkinter`; `ui/app.py` orchestrates scraping, caching, export and config I/O |

The single most important structural fact: **the project cannot be installed off Windows**, so no
test can ever run in CI as things stand.

```
$ uv sync
error: Distribution `pywin32==311` can't be installed because it doesn't have a
source distribution or wheel for the current platform
```

Everything else in this plan is downstream of fixing that.

---

## 2. Findings

### 2.1 Correctness bugs (confirmed by reading the code)

These are behaviour bugs, not style. Each becomes a failing test before it is fixed.

**B1 — VBA button-position code is never emitted.** `excel/macros/macros.py`
builds `self.code` with `{"\n".join(self.position_codes)}` **inside `__init__`**, but
`position_codes` is only populated later by `Button.create()` →
`macro.add_position_code(...)` (`excel/macros/button.py:33`). By the time
`spreadsheet.inject` reads `button.macro.code`, the string was already interpolated with an
empty list. The save/restore logic for button geometry is dead code in every generated
workbook.

**B2 — "Позиций ниже всех" is silently always wrong.**
`excel/formatters/analysis_formatter.py:29` uses
`self.data.get(competitor, {}).get(item, float('-inf'))`. When a competitor does not stock the
item, the comparison becomes `price < -inf`, which is always `False`, so `all(...)` fails and the
item is not counted. The intended "ignore missing" sentinel is `+inf`. Any item not stocked by
every single competitor is dropped from the metric.

**B3 — `App.config` shadows a Tk method.** `ui/app.py:36` assigns
`self.config = "config.json"`. `config` is `tkinter.Misc.config` (the alias of `configure`). Any
present or future `self.config(...)` call on the window raises `TypeError: 'str' object is not
callable`. Latent landmine.

**B4 — Empty profile crashes the parse.** `core/parser_engine.py:39` calls
`Pool(len(codes))`; with no entries this raises `ValueError: Number of processes must be at
least 1`. The exception escapes on a worker thread, so the progress bar spins forever.

**B5 — Tk is driven from a worker thread.** `App.start` runs on a `Thread` and calls
`CTkMessagebox(...)` (`ui/app.py:93`) and `self.done()` → `self.progress.stop()` /
`grid_forget()` / `os.startfile`. Tkinter is not thread-safe; this is an intermittent-crash
generator. All UI mutation must be marshalled back with `self.after(0, ...)`.

**B6 — `done(False)` iterates an error list nothing ever fills.** `ParserEngine.errors`
is initialised and never appended to; `ParserEngine.__slots__` also declares `callback`,
which is never assigned (an `AttributeError` waiting for a reader). The failure path in
`ui/app.py:114-116` therefore reports nothing. Also, the `try` in `App.click` only wraps
*starting* the thread, so real scraping failures are never caught at all.

**B7 — HTML parsing misaligns on any structural drift.** `core/parser_engine.py:21-27`
`zip()`s three independent CSS selections (`tooltip-info-header > a`, `span.form-title`,
`span.price-value`). If the counts differ — a promo row, a missing tooltip — `zip` truncates
and every subsequent name is paired with the wrong price. Silent data corruption, no warning.

**B8 — Price cleanup uses character-set strips.** `.rstrip(" р.").lstrip("от ")` removes
*any* of those characters from either end, not the literal affix. Should be
`removeprefix`/`removesuffix`. `float(price)` is also unguarded — one malformed cell aborts the
whole pharmacy.

**B9 — Sentinel strings mixed into a float column.** `data_formatter.py:31-36` appends the
string `"Нет"` into the same list as floats, then calls `float(price1)` on values that may be
that sentinel. It survives only because of the ordering of the two guards. `codes` on line 22 is
recomputed from `self.data.keys()` and is always equal to `self.titles`, so the
`y != codes[0]` guard reads as if it compares different things but does not.

**B10 — Column handling breaks past 26 columns.** Both formatters iterate
`string.ascii_uppercase`, so column widths, conditional formatting and diff columns silently stop
applying at column `Z` — i.e. at ~13 pharmacies. `get_column_letter` is already imported next to it.

**B11 — Realtime process priority for network I/O.** `utils/request.py:19` and
`core/parser_engine.py:19` call `psutil.Process(...).nice(psutil.REALTIME_PRIORITY_CLASS)` on
every fetch and every parse. Realtime priority on Windows can starve system processes; for an
I/O-bound scraper it buys nothing. Remove.

**B12 — Off-by-one temp-file dance.** `excel/spreadsheet.py:78` calls
`remove(f"{i - 1}{target}")` with `i = 0` on the first iteration, i.e. `remove("-1data.xlsm")`.
Harmless today, but symptomatic: the export writes `0data.xlsm`, `1data.xlsm`, … then
`os.rename`s the last one over `target` with no overwrite handling, and opens/quits a **whole
Excel process per sheet**.

**B14 — Duplicate pharmacy names corrupt the sheet.** Both the header builder in
`data_formatter.py` and the loop bound test `x != self.titles[-1]` compare *by value*, not by
index. Two pharmacies sharing a name make the header stop a column early while the data rows
still carry every price, so header and data misalign for the rest of the sheet. The same
duplicate also silently collapses in `dict(zip(titles, parse_res))` in `ParserEngine.process`.
Found while rewriting the header construction in phase 0; fix belongs with the domain model in
phase 1 (pharmacies need a stable identity separate from their display name).

**B15 — `settings.title` is dead configuration.** It is loaded, validated and written back,
but nothing reads it — verified by grep across the whole package. A real user's config sets it to
`"Анализ"`, so it plainly *looks* meaningful. Either wire it up (the natural home is a caption on
the analysis sheet, or the workbook's document title) or drop it; silently ignoring a value the
user set is the worst of the three options. Deferred to phase 3, which owns the sheet layout.

**B13 — `Request.url` is configured but ignored.** `utils/request.py` hardcodes
`HTTPSConnection("tabletka.by")` and `/ajax-request/reload-pharmacy-price/` while the `url` field
from `config.json` is loaded and never read. Changing the endpoint in config does nothing. There
are also no timeouts, no retries, no status-code checks, and no `try/except` around
`json.loads` — the app hangs or explodes on any network hiccup.

### 2.2 Architecture

**A1 — `ui/app.py` is a god object.** In 130 lines it owns: config load, config save, `Settings`
construction, `ParserEngine` construction, profile lifecycle, thread management, cache
read/write, export orchestration, and shelling out to `os.startfile`. None of it can be exercised
without a display server.

**A2 — There is no domain layer.** The entire model is
`DataType = Dict[str, Dict[str, float]]`. Comparison rules, "cheaper than all competitors",
assortment counts and percentage deltas live inside worksheet-writing loops
(`analysis_formatter.py:22-53` is a single 30-line nested literal). This is the main reason
nothing is testable: there is no function that takes prices and returns numbers.

**A3 — `utils/` is a junk drawer.** It holds a type alias, two enums, an HTTP client, a settings
dataclass, filesystem helpers **and a CustomTkinter widget factory**. Because
`utils/__init__.py` re-exports `create_custom_entry`, `import utils` imports the GUI toolkit —
so `core.parser_engine` transitively depends on CustomTkinter. Layering is inverted.

**A4 — Windows lock-in is unmarked and total.** `pythoncom` / `win32com` / `win32api` are
imported at module scope in `excel/spreadsheet.py`, so `import excel` — and therefore
`import ui` — fails on any non-Windows machine. Plus `os.startfile`, a hardcoded `"\\"` path
separator (`ui/app.py:113`), `LOCALAPPDATA`, and `psutil.REALTIME_PRIORITY_CLASS`. The
`.xlsx` half of the export needs none of this; only the macro buttons do.

**A5 — Wrong concurrency primitive.** `multiprocessing.Pool(len(codes))` spawns one OS process
per pharmacy for a purely I/O-bound job, pickling `ParserEngine` to each. `aiohttp` and
`requests` are both declared dependencies and **neither is used** — the code hand-rolls
`http.client.HTTPSConnection` instead.

**A6 — Config handling is unvalidated and lossy.** `Settings(**loaded["settings"])` on a
`__slots__` dataclass raises an opaque `TypeError` on any missing or extra key; a missing
`config.json` is an unhandled `FileNotFoundError` at startup with no guidance. Profile *names*
from the config file are discarded and rewritten as `Profile 1..N` on save (`ui/app.py:126`), so
user-chosen names cannot survive. The save is non-atomic — an interrupted write truncates the
file, taking the session cookies with it. Field names are `camelCase` in a Python dataclass.

**A7 — Cache is a hardcoded side effect.** `data.json` in the CWD, written unconditionally on
every successful parse, never invalidated, with no notion of staleness or which profile it
belongs to. The checkbox is read from the worker thread.

**A8 — `__slots__` is cargo-culted.** Declared on `App` (a Tk subclass, which has a `__dict__`
regardless — and `cache_checkbox` is missing from the list anyway), on dataclasses, and on
classes that also carry class-level attributes. It buys nothing here and misleads readers.

### 2.3 Packaging & tooling

**P1** — `pyproject.toml` lists ~13 transitive dependencies as direct ones (`altgraph`, `pefile`,
`certifi`, `idna`, `soupsieve`, `et-xmlfile`, `charset-normalizer`, `urllib3`, `packaging`,
`setuptools`, `pywin32-ctypes`, `darkdetect`, `typing-extensions`).
**P2** — `pywin32` has no environment marker → the hard install failure shown above.
**P3** — `pandas` + `numpy` are pulled in for a single `numpy.mean()` call.
**P4** — `pyinstaller` is a build tool sitting in runtime dependencies; there is no dev group.
**P5** — No `[build-system]`, no `[project.scripts]`, `description = "Add your description here"`.
**P6** — No ruff, no mypy, no pytest, no pre-commit, no CI workflow, no `.editorconfig`.
**P7** — `requires-python = ">=3.13"` is stricter than anything the code needs (nested f-string
quotes are 3.12+); it needlessly narrows the contributor and CI matrix.
**P8** — Undocumented hard requirement: VBA injection needs Excel's *"Trust access to the VBA
project object model"* enabled. A new user hits an opaque COM error with no hint in the README.
**P9** — Session cookies and the `_csrf` token live in plaintext `config.json` in the repo root.
It is covered by the blanket `*.json` in `.gitignore`, which also silently ignores any other JSON
a contributor adds — including fixtures. Worth narrowing and documenting.

---

## 3. Target architecture

A ports-and-adapters split. The rule to hold: **the domain layer imports nothing from
`openpyxl`, `customtkinter`, `win32com`, or the network.**

```
src/pharmparser/
├── __main__.py              # python -m pharmparser
├── cli.py                   # headless entry point — the CI-testable path
├── platform_.py             # is_windows(), open_file(), capability probes
│
├── config/
│   ├── models.py            # AppConfig / Settings / Profile / RequestConfig  (validated)
│   ├── loader.py            # load, atomic save, defaults, first-run bootstrap, migration
│   └── paths.py             # platformdirs-based location, no CWD assumptions
│
├── domain/                  # ← pure. no I/O, no third-party frameworks.
│   ├── models.py            # Pharmacy, DrugPrice, PriceTable, Comparison
│   └── analysis.py          # deltas, percentages, assortment/cheapest/unique metrics
│
├── scraping/
│   ├── client.py            # TabletkaClient: async, timeouts, retries, pagination
│   ├── parser.py            # parse_page(html) -> list[DrugPrice]   — pure
│   ├── service.py           # bounded-concurrency fan-out -> PriceTable
│   └── protocols.py         # PriceSource protocol (fakes live in tests)
│
├── export/
│   ├── protocols.py         # Exporter protocol
│   ├── grids.py             # build_data_grid / build_percent_grid / build_analysis_grid -> Grid
│   ├── xlsx_writer.py       # openpyxl only — cross-platform, always runs
│   └── vba/                 # Windows-only macro injection, optional post-step
│
└── ui/
    ├── controller.py        # UI-agnostic: owns state + use cases, testable headless
    ├── app.py               # thin: widgets + wiring only
    └── widgets/
tests/
├── unit/                    # domain, grids, parser, config
├── integration/             # xlsx round-trip via openpyxl read-back
└── fixtures/                # saved tabletka.by HTML, golden grids
```

Five design moves carry the refactor:

1. **Grids, not worksheets.** Sheet content becomes `list[list[Cell]]` produced by pure
   functions. Tests assert on lists — no Excel, no COM, no display. This alone makes ~all of
   the current formatter logic testable.
2. **VBA becomes optional.** `openpyxl` always writes a valid `.xlsx`. Macro injection is a
   separate post-step that is skipped with a logged warning when COM is unavailable. The app
   becomes runnable and testable on Linux/macOS; Windows users lose nothing.
3. **Async replaces `multiprocessing`.** One `aiohttp` session, a `Semaphore` for bounded
   concurrency, real timeouts and retries with backoff. Drops the process pool, the pickling,
   and the realtime-priority hack.
4. **A `Controller` between UI and use cases.** Holds profiles/state and exposes
   `parse()`, `export()`, `save_config()`. `App` becomes widget layout plus callbacks, and
   every UI update crosses back to the main thread through `self.after(0, ...)`.
5. **Validated config with named profiles.** Explicit parsing with actionable error messages,
   `snake_case` fields, atomic writes, and profile names preserved as the user set them.

---

## 4. Phased execution

Each phase ends with CI green and is independently mergeable. Phase 0 is a hard prerequisite
for everything else.

### Phase 0 — Make the project buildable and testable *(prerequisite)* — **done**
- Add `sys_platform == 'win32'` markers to `pywin32`; prune the ~13 transitive deps; move
  `pyinstaller` to a dev group. Drop `pandas`; replace the lone `numpy.mean` with `statistics.mean`.
- Add `[build-system]`, `[project.scripts]`, a real `description`; relax `requires-python` to `>=3.12`.
- Move to a `src/` layout; add `pytest` + `pytest-cov`, `ruff`, `mypy` config.
- Add a GitHub Actions workflow (lint → typecheck → test) and `pre-commit`.
- **Exit criteria:** `uv sync` succeeds on Linux; `pytest` runs; CI is green.

### Phase 1 — Extract the domain, fix the logic bugs — **done**
- Introduce `domain/models.py` and `domain/analysis.py`; lift every comparison and metric out of
  the formatters into pure functions.
- Write failing tests first for **B2, B9, B10**, then fix.
- Replace `DataType` with real types throughout.
- **Exit criteria:** analysis metrics fully unit-tested, including the missing-item cases.

**Outcome.** `domain/models.py` (`Pharmacy`, `PriceTable`) and `domain/analysis.py`
(`comparison_rows`, `summarise`, and friends) are pure — no openpyxl, no Tk, no COM, no
network — and carry every comparison rule that used to live inside worksheet loops. The
formatters became thin renderers over them. B2, B9 and B10 are fixed and covered by
regression tests; B14 is fixed in the domain (identity is `Pharmacy.id`, and the header is
built by index) with the legacy name-keyed adapter now *refusing* duplicate names instead of
silently dropping one — phase 2 removes the adapter by threading real ids through the scraper.

**One visible output change.** Fixing B9 means a difference cell for an item a pharmacy does
not stock is now left **blank** rather than written as `0`. A `0` was indistinguishable from
two prices that genuinely match. The VBA `">0"` filter excludes blanks and zeroes alike, so
filtering behaviour is unchanged; blanks sort last rather than among the zeroes.

### Phase 2 — Rewrite the scraping layer — **done**
- Split `client` (HTTP) / `parser` (HTML→models) / `service` (fan-out) apart.
- Async `aiohttp` with timeouts, retries + backoff, status checks, bounded concurrency.
- Fix **B7** (parse per result row instead of zipping parallel selections, and warn on
  mismatch), **B8** (`removeprefix`/`removesuffix`, guarded float), **B11** (delete the priority
  calls), **B13** (honour the configured URL), **B4** (empty-entry guard), **B6** (real error
  propagation).
- Save 2–3 real response pages into `tests/fixtures/` and test the parser against them.
- **Exit criteria:** parser is 100 % covered from fixtures; the network layer is faked in tests
  behind `PriceSource`.

**Outcome.** `scraping/` splits into `parser` (pure HTML → prices), `client` (async aiohttp with
timeouts, retries and backoff, validating the JSON envelope with pydantic) and `service`
(bounded-concurrency fan-out to a `PriceTable`). `multiprocessing.Pool` and the realtime-priority
calls are gone. B4, B6, B7, B8, B11 and B13 are fixed with regression tests.

**Validated against a real config.** `tests/fixtures/real_world_config.json` is an actual user
configuration with the credentials redacted: six profiles, Cyrillic and Latin pharmacy names,
names with trailing and doubled spaces, a pharmacy id reused across profiles, a full browser
header set (lowercase `host`, explicit `Content-Type` alongside form-encoded data) and non-default
column widths. It loads, round-trips byte-for-byte, scrapes and exports —
`tests/integration/test_real_world_config.py` keeps it that way.

**Fixtures are synthetic.** `tests/fixtures/*.html` were reconstructed from the CSS selectors the
old engine used, not captured from tabletka.by. They pin the parser's contract — row scoping,
price cleanup, tolerance of drift — but cannot prove the selectors still match the live site.
Replacing them with real captured responses is the single highest-value follow-up.

### Phase 3 — Rewrite the export layer
- Pure grid builders + golden-file tests.
- Single `openpyxl` writer; conditional formatting and column widths driven by
  `get_column_letter`, not `ascii_uppercase`.
- Isolate COM/VBA behind a capability check; fix **B1** (build macro source *after* buttons
  register their position code) and **B12** (write to a `tempfile` directory, one Excel process
  for the whole workbook, atomic replace at the end).
- **Exit criteria:** a full `.xlsx` export runs and is verified by read-back in CI on Linux.

### Phase 4 — Config and persistence — **done** (pulled forward)
- Validated models, actionable errors, first-run bootstrap from the example, atomic save,
  `snake_case`, preserved profile names (**A6**).
- Cache becomes explicit: opt-in, per-profile, timestamped, in a proper cache dir (**A7**).
- **Exit criteria:** malformed/missing/partial config all covered by tests.

**Outcome.** Pulled forward from its planned slot because the scraping rewrite needed a validated
`RequestConfig` and would otherwise have been built twice. `config/models.py` is pydantic:
snake_case internally, camelCase on disk via an alias generator, so **existing `config.json` files
load unchanged**. Colours, widths, pharmacy URLs and profile/pharmacy name uniqueness are all
validated, and `ConfigError` reports the offending field instead of a bare `TypeError`. Saves are
atomic and profile names are preserved (A6). Caches are per profile and versioned (A7).

`pydantic-settings` supplies `EnvOverrides`, so credentials can stay out of the file entirely:
`PHARMPARSER_COOKIE`, `PHARMPARSER_CSRF` and `PHARMPARSER_FILE_NAME` (also read from `.env`)
take precedence over `config.json`.

### Phase 5 — Thin the UI
- Introduce `Controller`; reduce `App` to layout + callbacks.
- Fix **B3** (rename `self.config`), **B5** (all Tk mutation via `after`), and surface real errors
  to the user.
- Remove the cargo-culted `__slots__` (**A8**); move `create_custom_entry` out of `utils` into
  `ui/widgets` (**A3**).
- **Exit criteria:** `Controller` is exercised headless; `App` contains no business logic.

### Phase 6 — Docs and release
- README: architecture, the VBA-trust prerequisite (**P8**), cross-platform behaviour, contributor
  setup. `CONTRIBUTING.md`. Narrow the `*.json` ignore rule and document credential handling (**P9**).
- Structured logging with a rotating file handler.
- Verify the PyInstaller build still produces a working Windows binary.

---

## 5. Decisions needed before Phase 3

1. **Must the Excel output stay `.xlsm` with VBA buttons?** The plan keeps them on Windows while
   making the `.xlsx` path work everywhere. If the buttons are negotiable, dropping VBA removes
   `pywin32`, COM, the temp-file dance and the trust-settings prerequisite in one stroke —
   roughly a third of the remaining complexity. Native Excel table filters plus a frozen header
   row cover most of what the buttons do.
2. **Is a headless CLI wanted alongside the GUI?** Recommended: it is the cheapest way to keep
   the whole pipeline honest in CI, and it costs about 50 lines once the controller exists.
3. **Should the UI stay Russian-only?** Labels are currently hardcoded Russian string literals
   inside the formatters. If localisation is ever wanted, Phase 3 is the moment to route them
   through a message catalogue rather than retrofitting later.

---

## 6. Decisions taken (2026-08-18)

1. **VBA buttons are essential — keep as-is.** The `.xlsm` + COM path stays first-class on
   Windows. Consequence for the refactor: COM imports must become *lazy* rather than
   module-scope, so that the package is importable — and therefore testable — on Linux even
   though the macro step itself only runs on Windows. Cross-platform support is a test-suite
   and CI concern, not a user-facing feature. Phase 3 keeps the VBA layer and fixes B1/B12
   inside it instead of deleting it.
2. **A headless CLI will be added.** `pharmparser.cli` becomes the end-to-end path exercised in
   CI, and makes the scrape/export pipeline scriptable. *Done:* `pharmparser-cli`, covered by
   `tests/integration/test_cli.py`, which runs the whole pipeline against a faked endpoint.
3. **Execution starts at Phase 0.**
4. **Python 3.14** is the target runtime (added 2026-08-18). This forced an `lxml` floor of 6.0.1,
   the first release shipping cp314 wheels.
5. **pydantic and pydantic-settings** are adopted for configuration (added 2026-08-18), which
   pulled phase 4 forward ahead of phase 3.
