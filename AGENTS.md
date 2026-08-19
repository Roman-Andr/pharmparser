# Repository Guidelines

## Project Structure & Module Organization

Application code lives in `src/pharmparser/`. Keep business rules in `domain/`, HTTP and parsing code in `scraping/`, workbook generation in `export/`, and CustomTkinter views in `ui/`. `controller.py` coordinates application behavior without importing UI toolkits. PyInstaller entry points are in `packaging/`; developer utilities are in `tools/`. Tests mirror the product under `tests/unit/` and `tests/integration/`, with captured pages, redacted configuration, and golden data in `tests/fixtures/`.

Preserve the dependency boundaries documented in `CONTRIBUTING.md`: domain code performs no I/O, `export/grids.py` does not import OpenPyXL, and Windows-only `pythoncom`/`win32com` imports stay inside functions.

## Build, Test, and Development Commands

- `uv sync --all-groups` installs runtime, development, and build dependencies for Python 3.14+.
- `uv run pharmparser` starts the GUI; `uv run pharmparser-cli --help` shows CLI options.
- `uv run ruff check .` runs lint and import-order checks.
- `uv run mypy` type-checks `src/` and `tests/`.
- `uv run pytest` runs the suite. Use `xvfb-run -a uv run pytest` on headless Linux for GUI tests.
- `uv run pyinstaller --noconfirm pharmparser.spec` builds GUI and CLI binaries in `dist/`.

Run lint, type checking, and tests before committing.

## Coding Style & Naming Conventions

Use four-space indentation, type annotations, and a 120-character line limit. Ruff enforces pycodestyle, Pyflakes, import sorting, pyupgrade, Bugbear, comprehensions, simplifications, and Ruff-specific rules. Use `snake_case` for modules/functions, `PascalCase` for classes, and descriptive test names such as `test_a_checksum_mismatch_is_refused`. Cyrillic UI and workbook text is intentional.

## Testing Guidelines

Pytest and pytest-asyncio are configured in `pyproject.toml`; CI also records branch coverage. Add a regression test before fixing a bug. Unit tests should avoid I/O. Use the local HTTP fixture for network behavior and `FakeExcel` for COM behavior. Review golden-file diffs as user-visible workbook changes.

## Commit & Pull Request Guidelines

Use Conventional Commits and matching branch names, for example `fix/update-restart` or `feat/new-export`. `fix:` produces a patch release; `feat:` produces a minor release through release-please. PRs should explain impact and root cause, list validation commands, link relevant issues, and include screenshots for visible UI changes. Keep unrelated files out of commits and ensure CI passes before merge.

## Security & Configuration

Never commit `config.json`, cookies, CSRF tokens, or unredacted captured data. Inspect `git status` before staging and keep `tests/fixtures/real_world_config.json` redacted.
