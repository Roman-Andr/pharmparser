# PharmParser

> Desktop application for parsing and comparing drug prices from [tabletka.by](https://tabletka.by)
> pharmacies. Exports results to a formatted Excel workbook.

---

## Features

- Parse prices from many pharmacies at once, concurrently
- Compare every price against a reference pharmacy, with colour-coded differences
- React desktop interface in Russian, with light and dark themes
- Explicit reference pharmacy, per-pharmacy progress, cancellation and partial retries
- Local SQLite history with pinning, retention, trends and offline re-export
- `.xlsm` reports with dynamic Excel Tables and one-click VBA actions; `.xlsx` is optional
- Windows WebView2 host with browser fallback, plus a cross-platform headless CLI

---

## Requirements

- Python 3.14+
- [`uv`](https://github.com/astral-sh/uv)
- [`Bun`](https://bun.sh/) for frontend development only

The package installs and runs on Linux, macOS and Windows, macro buttons included:
the `.xlsm` is assembled in Python, so no Excel install is involved.

---

## Installation

```bash
git clone https://github.com/Roman-Andr/pharmparser.git
cd pharmparser
uv sync --all-groups
cd frontend && bun install && bun run build && cd ..
```

> The `.xlsm` is built in Python — no Excel needed. Excel still has to *trust* the
> macros when you open the file: allow content when prompted, or unblock the file in
> its file properties. Pass `--use-excel` to have Excel itself write the workbook
> over COM instead (Windows only, and needs the Trust Center's "Trust access to the
> VBA project object model").

---

## First launch and configuration

Run `uv run pharmparser`. The Russian onboarding flow configures the theme,
Cookie/CSRF, profile, explicit reference pharmacy, connection and output directory.

| Data | Windows | Linux |
| --- | --- | --- |
| Settings and profiles | `%APPDATA%/PharmParser/settings.json` | `$XDG_CONFIG_HOME/PharmParser/settings.json` |
| History and prices | `%LOCALAPPDATA%/PharmParser/history.sqlite3` | `$XDG_DATA_HOME/PharmParser/history.sqlite3` |
| Logs | `%LOCALAPPDATA%/PharmParser/logs/` | `$XDG_DATA_HOME/PharmParser/logs/` |
| Cookie and CSRF | Windows Credential Manager | system keyring, environment, or warned `0600` fallback |

Secrets are never written to settings, SQLite, reports, frontend responses or logs.
`PHARMPARSER_COOKIE` and `PHARMPARSER_CSRF` remain supported for headless systems.

### Legacy CLI configuration

The CLI keeps accepting the transition-release `config.json` format:

```bash
cp config.json.example config.json
```

```jsonc
{
  "profiles": {
    "Profile 1": {
      "Аптека 1": "https://tabletka.by/pharmacies/****"  // pharmacy page URL
    }
  },
  "settings": {
    "green": "19CF1F",     // fill for a competitor price above the reference
    "red": "E81737",       // fill for a competitor price below the reference
    "title": "Анализ",     // names the summary sheet and the workbook
    "fileName": "data.xlsx",
    "colWidth": 50,        // width of the item-name column
    "cellWidth": 15,       // width of a price column
    "diffWidth": 10        // width of a "Разница" column
  },
  "request": {
    "url": "https://tabletka.by/ajax-request/reload-pharmacy-price",
    "headers": {
      "Cookie": "PHPSESSID=...; _csrf=...; regionId=...; lim-result=5000"
    },
    "data": {
      "sort": "name",
      "sort_type": "asc",
      "str": "",
      "_csrf": "..."
    }
  }
}
```

The desktop app stores an explicit reference pharmacy. During legacy migration only,
the first pharmacy is imported as the reference. Values are validated on load, and a bad one names itself — a malformed colour, a
non-positive width, a pharmacy URL without a numeric id, a `title` Excel would reject as a
sheet name, or two pharmacies in one profile sharing a display name.

### Getting your session cookies

1. Open [tabletka.by](https://tabletka.by) in your browser
2. Open DevTools → **Network**
3. Load any pharmacy's price page
4. Copy the `Cookie` request header and the `_csrf` value from the request payload
5. Paste them into the desktop «Настройки» screen (or legacy `config.json` for CLI use)

The cookie must include `lim-result=5000` — the app rewrites that value to page through
results. Cookies expire; when scraping starts failing, this is the first thing to refresh.

### Headless credentials

Credentials can come from the environment (or a `.env` file) instead, and take precedence
over `config.json`:

| Variable | Overrides |
| --- | --- |
| `PHARMPARSER_COOKIE` | `request.headers.Cookie` |
| `PHARMPARSER_CSRF` | `request.data._csrf` |
| `PHARMPARSER_FILE_NAME` | `settings.fileName` |

`config.json`, `.env` and the cache files are all git-ignored by name. The ignore rule is
deliberately narrow rather than a blanket `*.json`, which used to swallow test fixtures
too — so check what you are committing if you add JSON of your own.

---

## Usage

```bash
uv run pharmparser        # GUI (same as: uv run python -m pharmparser)
uv run pharmparser-cli    # headless
```

In the desktop app, finish onboarding and press «Сформировать отчет». A single active
job shows separate progress for every pharmacy and may be cancelled. A partial result
is allowed only when the reference and at least one competitor succeeded; failed
pharmacies can be retried for 30 minutes while fresh results are reused. Reports never
open Excel without an explicit click.

The CLI runs the same pipeline:

```bash
uv run pharmparser-cli --profile "Основной" --output report.xlsx
uv run pharmparser-cli --macros    # also add the VBA buttons (any platform)
uv run pharmparser-cli --cache     # reuse this profile's cached prices
uv run pharmparser-cli --help
```

Every run also appends to a rotating `pharmparser.log` beside the config file, which is
the thing to look at — or ask a user for — when something goes wrong.

### What is in the desktop workbook

| Sheet | Contents |
| --- | --- |
| `Обзор` | Completeness, reference pharmacy, KPI, assortment and warnings |
| `Сравнение` | Structured product fields and price/difference columns per competitor |
| `Изменения` | Previous/current prices, changes, new and disappeared products |
| `История` | Full-run assortment and median price-index trends with a chart |
| `Проблемы` | Failed pharmacies, collisions, warnings and exact timestamps |

An item is labelled `name, pack, manufacturer` — the manufacturer is on the end so the
column still reads and sorts by drug name. It is part of the label because the same drug
and pack from two makers sells at two prices, and without it one of them was silently
dropped (about 1 % of rows).

A blank difference cell means the comparison is undefined — one of the two pharmacies does
not stock the item. That is distinct from a difference of `0`, which means the prices match.

XLSM is the desktop default. Its buttons operate on dynamic Excel Tables rather than
fixed row limits: «Наша цена ниже всех», «Есть конкурент дешевле», filter reset and
per-competitor sorting. XLSX keeps native table filters but has no buttons.

---

## Architecture

The rule that holds the layout together: **the domain layer imports nothing from
`openpyxl`, FastAPI, React, `win32com`, or the network.**

```bash
pharmparser/
├── src/pharmparser/
│   ├── __main__.py            # GUI entry point
│   ├── cli.py                 # headless entry point
│   ├── application/           # profiles, credentials, runs, history and report use cases
│   ├── controller.py          # compatibility adapter for transition entry points
│   ├── domain/                # pure model + analysis (no I/O, no frameworks)
│   │   ├── models.py          # Pharmacy, PriceTable
│   │   └── analysis.py        # comparisons, market summary
│   ├── config/                # pydantic schema, loader, paths, env overrides
│   ├── scraping/              # async client, pure HTML parser, fan-out service
│   ├── export/
│   │   ├── grids.py           # pure sheet builders (content + layout)
│   │   ├── xlsx_writer.py     # the only module that knows openpyxl
│   │   ├── protocols.py       # the Exporter contract
│   │   └── vba/               # macro buttons: VBA compiler, .xlsm packer, COM injector
│   ├── web/                   # protected FastAPI API, React assets and pywebview host
│   ├── logging_.py            # console + rotating file logging
│   ├── platform_.py           # OS capability probes
├── frontend/                  # React, TypeScript and Vite; locked with Bun
├── packaging/                 # PyInstaller entry points
├── tests/                     # unit, integration, fixtures
├── docs/REFACTOR_PLAN.md      # findings, target architecture, phase outcomes
├── pharmparser.spec           # PyInstaller build definition
└── config.json.example
```

Four ideas carry the design:

1. **Grids, not worksheets.** A sheet's content and layout are built as data by pure
   functions, so the report is testable without Excel. `xlsx_writer` is the only module
   that touches openpyxl.
2. **VBA is an optional post-step.** openpyxl always writes a valid `.xlsx`; the macro
   buttons are injected afterwards by driving Excel over COM. `pythoncom` and `win32com`
   are imported lazily, so the package stays importable — and testable — everywhere.
3. **Application services instead of UI-owned behavior.** FastAPI and the compatibility
   entry points call testable services; the local API only adapts these use cases.
4. **Local API hardening.** A 256-bit token lives in the URL fragment, every API request
   uses Bearer auth, and Host/Origin/CSP checks keep the random-port loopback server private.
5. **Immutable history.** SQLite stores money as integer kopecks, snapshots profiles and
   retains structured products. Completed runs alone become trend baselines.

See [`docs/REFACTOR_PLAN.md`](docs/REFACTOR_PLAN.md) for the findings behind all of this
and what each phase changed.

---

## Releases and updating

Tagged versions are published to [GitHub Releases](https://github.com/Roman-Andr/pharmparser/releases)
with a Windows and a Linux binary of both the GUI and the CLI, plus a `SHA256SUMS`
file and a build provenance attestation.

### Cutting a release

Version numbers are not chosen by hand. Every push to `main` updates a standing
**Release PR** carrying the next version and the changelog earned since the last
release, worked out from the commit prefixes (`fix:` bumps the patch, `feat:` the
minor — see [CONTRIBUTING](CONTRIBUTING.md#commit-messages)).

Releasing is merging that PR. That tags the version, writes `CHANGELOG.md`, builds
the binaries and attaches them.

So: nothing to remember, and nothing published until you merge. Merges to `main`
that are not the Release PR publish nothing — CI uploads binaries as workflow
artifacts for testing instead. Releases are what the auto-updater follows, so they
stay deliberate.

`workflow_dispatch` on the Release workflow rebuilds and re-attaches assets for an
existing tag, if a build ever needs redoing.

### Updating

The packaged app checks for a newer release shortly after it starts and offers it;
it never replaces itself without being asked. The download is verified against the
published `SHA256SUMS` before anything is run, and any release URL outside GitHub is
refused. Running from a source checkout skips the check entirely.

The headless binary can be asked directly, and installs nothing:

```bash
pharmparser-cli --check-update
```

> The binaries are **not code-signed**, so Windows SmartScreen warns on first run.
> The checksum protects a download from corruption or tampering in transit; the trust
> root is still GitHub account security, not a certificate. Signing needs a paid
> certificate.

---

## Troubleshooting

**"the session looks expired"** — the cookies in `config.json` have a limited life,
and this is the failure you will hit most often. Refresh the `Cookie` header and the
`_csrf` value from DevTools (Network tab, any request the prices page makes) and run
again. The app fails fast on this rather than retrying, so it takes about a second
to find out.

**"check the pharmacy URL"** — the endpoint answered 404 for that pharmacy id. The
URL in the profile must end with the numeric id, as in
`https://tabletka.by/pharmacies/3563`.

**The report opened without the sort and filter buttons** — the workbook was written
as `.xlsx` rather than `.xlsm`. Pass `--macros` on the CLI, and check the log for a
line about the VBA project.

**Excel warns about macros on first open** — expected. The workbook carries a VBA
project, and Excel blocks macros in files downloaded from the internet until you
allow the content, or unblock the file in its properties.

**"Parse workers unavailable"** — parsing fell back to a single process, so the run
is slower but correct. It happens when the app is embedded in a script with no
`if __name__ == "__main__":` guard, because worker processes re-import the entry
module.

---

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md). The short version:

```bash
uv sync --all-groups     # runtime + dev + build dependencies
uv run ruff check .      # lint
uv run mypy              # type check
uv run pytest            # tests
uv run pre-commit install
cd frontend
bun install
bun run generate:api     # refresh OpenAPI contract and generated TypeScript types
bun run test
bun run build
```

The GUI smoke tests need a display and skip without one. On a headless machine:

```bash
xvfb-run -a uv run pytest
```

### Building a binary

```bash
uv run pyinstaller pharmparser.spec
```

Produces `dist/pharmparser` (GUI, no console) and `dist/pharmparser-cli` (headless), both
single-file. PyInstaller 6.15 is the floor: earlier releases cap out at Python 3.13 and
their bootloader aborts with *Failed to allocate PyConfig structure* on a 3.14 build.

---

## Tech Stack

| Layer | Library |
| --- | --- |
| Desktop | React, TypeScript, Vite, pywebview/WebView2, FastAPI |
| HTTP | `aiohttp` |
| Parsing | `beautifulsoup4`, `lxml` |
| Config | `pydantic`, `pydantic-settings` |
| Export | `openpyxl`, `ms-ovba` (VBA project), `pywin32` (optional Windows COM) |
| Packaging | `pyinstaller` |

---

## License

[MIT](LICENSE)
