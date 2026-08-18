# PharmParser

> Desktop application for parsing and comparing drug prices from [tabletka.by](https://tabletka.by) pharmacies. Exports results to a formatted Excel spreadsheet.

---

## Features

- Parse prices from multiple pharmacies in one click
- Compare prices across pharmacies with color-coded highlighting
- Multi-profile support — save different pharmacy sets and switch between them
- Export results to `.xlsx` with customizable column widths and color settings
- Clean GUI built with CustomTkinter (supports light/dark/system theme)

---

## Requirements

- Python 3.14+
- [`uv`](https://github.com/astral-sh/uv) package manager

---

## Installation

```bash
git clone https://github.com/RomanAndr/pharmparser.git
cd pharmparser
uv sync
```

> **Windows only:** generating the `.xlsm` workbook drives Excel over COM and writes a VBA
> module, which requires Excel's *Trust access to the VBA project object model* setting.
> Enable it under **File → Options → Trust Center → Trust Center Settings → Macro Settings**.
> Without it the export fails with an opaque COM error.

---

## Configuration

Copy the example config and fill in your session credentials:

```bash
cp config.json.example config.json
```

`config.json` structure:

```jsonc
{
  "profiles": {
    "Profile 1": {
      "1": "https://tabletka.by/pharmacies/****"  // pharmacy page URL
    }
  },
  "settings": {
    "green": "19CF1F",     // color for lower price
    "red": "E81737",       // color for higher price
    "title": "My Report",
    "fileName": "data.xlsx",
    "colWidth": 50,
    "cellWidth": 15,
    "diffWidth": 10
  },
  "request": {
    "url": "https://tabletka.by/ajax-request/reload-pharmacy-price",
    "headers": {
      "Cookie": "PHPSESSID=...; _csrf=...; regionId=..."
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

### How to get your session cookies

1. Open [tabletka.by](https://tabletka.by) in your browser and log in
2. Open DevTools → **Network** tab
3. Make any request to the pharmacy prices page
4. Copy the `Cookie` header value and the `_csrf` token from the request payload
5. Paste them into `config.json`

You can use **Postman** or any HTTP client to verify the request works before running the app.

---

## Usage

```bash
uv run pharmparser            # GUI
uv run pharmparser-cli        # headless
```

(the GUI is equivalently `uv run python -m pharmparser`)

The CLI runs the same pipeline without a display:

```bash
uv run pharmparser-cli --profile "Основной" --output report.xlsx
uv run pharmparser-cli --macros        # also inject the VBA buttons (Windows)
uv run pharmparser-cli --help
```

### Keeping credentials out of the config file

Session credentials may be supplied by the environment (or a `.env` file) instead of
`config.json`, and take precedence over it:

| Variable | Overrides |
| --- | --- |
| `PHARMPARSER_COOKIE` | `request.headers.Cookie` |
| `PHARMPARSER_CSRF` | `request.data._csrf` |
| `PHARMPARSER_FILE_NAME` | `settings.fileName` |

1. Select or create a profile with your pharmacy URLs
2. Click **Parse** — the app fetches current prices
3. The result is saved as an `.xlsx` file defined in `fileName`

---

## Project Structure

```bash
pharmparser/
├── src/pharmparser/
│   ├── __main__.py            # GUI entry point
│   ├── cli.py                 # headless entry point
│   ├── domain/                # pure model + analysis (no I/O, no frameworks)
│   │   ├── models.py          # Pharmacy, PriceTable
│   │   └── analysis.py        # comparisons, market summary
│   ├── config/                # pydantic schema, loader, env overrides
│   ├── scraping/              # async client, pure HTML parser, fan-out service
│   ├── export/                # price table -> workbook
│   │   ├── grids.py           # pure sheet builders (content + layout)
│   │   ├── xlsx_writer.py     # the only module that knows openpyxl
│   │   └── vba/               # Windows-only macro buttons, imported lazily
│   ├── cache.py               # per-profile scrape cache
│   ├── platform_.py           # OS capability probes
│   └── ui/                    # CustomTkinter windows and widgets
├── tests/
│   ├── unit/
│   └── integration/           # workbook round-trip, no Excel required
├── docs/REFACTOR_PLAN.md      # in-progress restructuring plan
├── config.json.example
└── pyproject.toml
```

> The layout above is mid-refactor. See [`docs/REFACTOR_PLAN.md`](docs/REFACTOR_PLAN.md)
> for the target architecture and the phased plan to get there.

---

## Development

```bash
uv sync --all-groups     # runtime + dev + build dependencies
uv run pytest            # tests
uv run ruff check .      # lint
uv run mypy              # type check
uv run pre-commit install
```

The package installs and the test suite runs on Linux and macOS as well as Windows;
only the Excel/COM macro-injection step is Windows-specific.

Tests marked `xfail(strict=True)` pin known bugs that are documented in the refactor
plan and fixed in a later phase — if one starts passing unexpectedly, the suite fails,
so the fix cannot land silently.

---

## Tech Stack

| Layer | Library |
| --- | --- |
| GUI | [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) |
| HTTP | `aiohttp` |
| Parsing | `beautifulsoup4`, `lxml` |
| Config | `pydantic`, `pydantic-settings` |
| Export | `openpyxl`, `pywin32` (Windows COM/VBA) |
| Packaging | `pyinstaller` |

---

## License

[MIT](LICENSE)
