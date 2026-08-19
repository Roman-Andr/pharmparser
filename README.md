# PharmParser

> Desktop application for parsing and comparing drug prices from [tabletka.by](https://tabletka.by)
> pharmacies. Exports results to a formatted Excel workbook.

---

## Features

- Parse prices from many pharmacies at once, concurrently
- Compare every price against a reference pharmacy, with colour-coded differences
- Multi-profile support — save different pharmacy sets and switch between them
- Export to `.xlsx`, or to `.xlsm` with VBA sort/filter buttons — both on any platform
- A GUI built with CustomTkinter, and a headless CLI running the same pipeline

---

## Requirements

- Python 3.14+
- [`uv`](https://github.com/astral-sh/uv)

The package installs and runs on Linux, macOS and Windows, macro buttons included:
the `.xlsm` is assembled in Python, so no Excel install is involved.

---

## Installation

```bash
git clone https://github.com/Roman-Andr/pharmparser.git
cd pharmparser
uv sync
```

> The `.xlsm` is built in Python — no Excel needed. Excel still has to *trust* the
> macros when you open the file: allow content when prompted, or unblock the file in
> its file properties. Pass `--use-excel` to have Excel itself write the workbook
> over COM instead (Windows only, and needs the Trust Center's "Trust access to the
> VBA project object model").

---

## Configuration

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

The first pharmacy in a profile is the **reference**: every other one is compared against
it. Values are validated on load, and a bad one names itself — a malformed colour, a
non-positive width, a pharmacy URL without a numeric id, a `title` Excel would reject as a
sheet name, or two pharmacies in one profile sharing a display name.

### Getting your session cookies

1. Open [tabletka.by](https://tabletka.by) in your browser
2. Open DevTools → **Network**
3. Load any pharmacy's price page
4. Copy the `Cookie` request header and the `_csrf` value from the request payload
5. Paste them into `config.json`

The cookie must include `lim-result=5000` — the app rewrites that value to page through
results. Cookies expire; when scraping starts failing, this is the first thing to refresh.

### Keeping credentials out of the config file

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

In the GUI: pick or create a profile, fill in pharmacy names and URLs, press **Parse**.
The workbook opens when it is ready. Tick **Cache** to reuse the last scrape for that
profile instead of hitting the network — the cache is per profile and entirely opt-in.

The CLI runs the same pipeline:

```bash
uv run pharmparser-cli --profile "Основной" --output report.xlsx
uv run pharmparser-cli --macros    # also add the VBA buttons (any platform)
uv run pharmparser-cli --cache     # reuse this profile's cached prices
uv run pharmparser-cli --help
```

Every run also appends to a rotating `pharmparser.log` beside the config file, which is
the thing to look at — or ask a user for — when something goes wrong.

### What is in the workbook

| Sheet | Contents |
| --- | --- |
| `Данные` | One row per item: the reference price, then each competitor's price and the difference in roubles |
| `Проценты` | The same layout, with differences as a percentage of the reference price |
| `settings.title` | Assortment sizes, items cheapest everywhere, unique items, and a per-competitor breakdown |

An item is labelled `name, pack, manufacturer` — the manufacturer is on the end so the
column still reads and sorts by drug name. It is part of the label because the same drug
and pack from two makers sells at two prices, and without it one of them was silently
dropped (about 1 % of rows).

A blank difference cell means the comparison is undefined — one of the two pharmacies does
not stock the item. That is distinct from a difference of `0`, which means the prices match.

With `--macros` the two price sheets also carry buttons: **Apply Filters** / **Remove Filters**,
and an up/down pair over each difference column.

---

## Architecture

The rule that holds the layout together: **the domain layer imports nothing from
`openpyxl`, `customtkinter`, `win32com`, or the network.**

```bash
pharmparser/
├── src/pharmparser/
│   ├── __main__.py            # GUI entry point
│   ├── cli.py                 # headless entry point
│   ├── controller.py          # state + use cases, driven by both front ends
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
│   ├── cache.py               # per-profile scrape cache
│   ├── logging_.py            # console + rotating file logging
│   ├── platform_.py           # OS capability probes
│   └── ui/                    # CustomTkinter windows and widgets
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
3. **A controller between the front ends and the use cases.** Both the GUI and the CLI
   drive the same `Controller`, so they cannot drift apart, and it runs headless.
4. **Validated config.** Explicit models with actionable errors, atomic saves, and the
   on-disk format preserved exactly, so existing `config.json` files keep working.

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

`workflow_dispatch` on the Release workflow does two things: with a tag it rebuilds
and re-attaches that tag's assets, and with the tag left empty it just refreshes the
Release PR — useful if a run failed for a reason outside the workflow.

> Release-please opens that PR through GitHub Actions, which needs
> **Settings → Actions → General → Workflow permissions →
> "Allow GitHub Actions to create and approve pull requests"**. Without it the run
> fails at the very last step, having already computed the version and pushed the
> release branch.

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

## Development

See [CONTRIBUTING.md](CONTRIBUTING.md). The short version:

```bash
uv sync --all-groups     # runtime + dev + build dependencies
uv run ruff check .      # lint
uv run mypy              # type check
uv run pytest            # tests
uv run pre-commit install
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
| GUI | [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) |
| HTTP | `aiohttp` |
| Parsing | `beautifulsoup4`, `lxml` |
| Config | `pydantic`, `pydantic-settings` |
| Export | `openpyxl`, `ms-ovba` (VBA project), `pywin32` (optional Windows COM) |
| Packaging | `pyinstaller` |

---

## License

[MIT](LICENSE)
