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

- Python 3.13+
- [`uv`](https://github.com/astral-sh/uv) package manager

---

## Installation

```bash
git clone https://github.com/RomanAndr/pharmparser.git
cd pharmparser
uv sync
```

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
uv run main.py
```

1. Select or create a profile with your pharmacy URLs
2. Click **Parse** — the app fetches current prices
3. The result is saved as an `.xlsx` file defined in `fileName`

---

## Project Structure

```bash
pharmparser/
├── core/
│   └── parser_engine.py   # HTTP requests & HTML parsing logic
├── ui/
│   ├── app.py             # Main application window
│   ├── entry.py           # Input fields
│   ├── profile.py         # Profile management UI
│   └── profile_selector.py
├── utils/
│   ├── datatypes.py
│   ├── file_utils.py      # Excel export
│   ├── filter_criteria.py
│   ├── request.py         # Request builder
│   ├── settings.py        # Config loader
│   ├── sort_order.py
│   └── widgets.py         # Reusable UI components
├── excel/                 # Excel formatting helpers
├── main.py
├── config.json.example
└── pyproject.toml
```

---

## Tech Stack

| Layer | Library |
| --- | --- |
| GUI | [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) |
| HTTP | `requests`, `aiohttp` |
| Parsing | `beautifulsoup4`, `lxml` |
| Export | `openpyxl`, `pandas` |
| Packaging | `pyinstaller` |

---

## License

[MIT](LICENSE)
