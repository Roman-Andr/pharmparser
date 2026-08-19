# Contributing

## Setup

```bash
uv sync --all-groups     # runtime + dev + build dependencies
uv run pre-commit install
```

Python 3.14 or newer. Everything except the Excel/COM macro step runs on Linux,
macOS and Windows, so you do not need Windows to work on this.

## The checks

All three must pass before a commit:

```bash
uv run ruff check .
uv run mypy
uv run pytest
```

The GUI smoke tests need a display. They skip without one; on a headless machine:

```bash
xvfb-run -a uv run pytest
```

## Commit messages

[Conventional Commits](https://www.conventionalcommits.org/). This is not a style
preference: release-please reads the prefixes to decide the next version number and
to write the changelog, so the prefix on a commit is what ships it.

| prefix | effect on the version | appears in the changelog |
| --- | --- | --- |
| `fix:` | patch — `0.2.0` -> `0.2.1` | yes |
| `feat:` | minor — `0.2.1` -> `0.3.0` | yes |
| `feat!:`, `BREAKING CHANGE:` | minor while pre-1.0, major after | yes, highlighted |
| `perf:`, `refactor:`, `docs:`, `build:` | none | yes |
| `ci:`, `test:`, `chore:` | none | no |

A commit with no recognised prefix releases nothing and says nothing, so it is worth
getting right. Write the body for someone reading `git log` a year from now.

## Branches

`main` is what releases. `dev` is where dependency updates collect: Dependabot opens
its pull requests there rather than against `main`, so a week of routine bumps can be
reviewed and merged forward together instead of one at a time.

Feature work can go either way — straight to `main` for anything self-contained, via
`dev` when it should ride along with the dependency bumps. Merging `dev` into `main`
is what puts any of it into a release.

Name a working branch after the commit type it carries, so the branch says the same
thing the changelog will:

```
feat/excel-without-com      fix/stale-cookie-message      perf/lxml-parser
docs/troubleshooting        ci/branch-naming              refactor/export-layer
```

Only `main` and `dev` build on push. Everything else is covered by the pull request
trigger, so a branch gets its checks when it is proposed rather than twice over.

Two things about Dependabot that are GitHub's behaviour, not ours:

* `.github/dependabot.yml` is read from the default branch only, so it lives on
  `main` even though every PR it opens targets `dev`.
* **Security** updates ignore `target-branch`. They are raised against `main`, and
  they do not carry the commit-message prefixes from that file — so one arrives
  titled `Bump x from 1 to 2`, which release-please does not read as anything.
  Retitle it to `fix(deps): …` before merging if it should cut a release.

## Merging

The merge strategy decides what release-please sees, so it is a release decision
rather than a taste one: only commits that land on `main` are read, and a squash
collapses a whole branch into one.

**Rebase and merge** when the branch's commits are already one-idea-each with proper
prefixes. Every one of them then earns its own changelog line, and the highest type
among them sets the version bump.

**Squash and merge** when the branch has fixup and work-in-progress commits worth
hiding. Then the squash message is the *only* thing released, so it has to be a
proper Conventional Commit itself — GitHub defaults the subject to the PR title, so
the PR title becomes the release note. A title like "updates" releases nothing and
says nothing. If squashing becomes the norm here, add a PR-title lint
(`amannn/action-semantic-pull-request`) so that failure is caught in review rather
than discovered at release time.

Squashing does not fold the branch's commits into several changelog entries: they
end up in the squash commit's *body*, and release-please reads one type per commit,
from the subject.

The Release PR itself is a single commit either way, so merge it however you like.

## How the code is laid out

The rule that keeps this testable: **the domain layer imports nothing from
`openpyxl`, `customtkinter`, `win32com`, or the network.** In practice:

| If you are changing… | …the layer is | and it must not import |
| --- | --- | --- |
| a comparison or a metric | `domain/` | anything with I/O or a framework |
| what a sheet contains | `export/grids.py` | openpyxl |
| how a workbook is written | `export/xlsx_writer.py` | COM, the network |
| the macro buttons | `export/vba/` | anything at module scope from pywin32 |
| what the app *does* | `controller.py` | a toolkit |
| what the window *looks like* | `ui/` | business logic |

Two rules follow from that and are enforced by tests:

- `pythoncom` and `win32com` are imported **inside functions**, never at module
  scope, so the package stays importable on Linux.
- Importing `pharmparser.controller`, `pharmparser.cli` or `pharmparser.ui` must
  not load `tkinter`. Widget classes resolve lazily through `ui/__init__.py`.

## Testing conventions

- **Unit tests** get no I/O: the domain, the grids, the parser and the config are
  all pure, so test them as functions over data.
- **The network** is a real HTTP server on localhost (`tests/endpoint.py`), not a
  patched aiohttp. It answers from the request's own fields, so concurrency cannot
  make a test flaky.
- **Excel** is a recorder standing in for the COM object model (`tests/fakes.py`),
  which lets the Windows-only export path run in CI on any platform.
- **Page markup** for tests comes from `tests/pages.py`, whose template mirrors a
  response captured from the live site. A test keeps the two in step — see
  `tests/fixtures/README.md` before changing either.
- **Golden files** pin the exported sheets (`tests/fixtures/golden_grids.json`).
  A diff there is the diff a user would see in Excel: read it before regenerating.

Cover a bug with a failing test before fixing it, and say in the test *what* broke
rather than only what should happen — most tests here name the bug they pin.

## Credentials

Never commit `config.json`, a session cookie or a `_csrf` token. The ignore rules
name those files specifically rather than ignoring all JSON, so check `git status`
if you add JSON of your own. `tests/fixtures/real_world_config.json` is a real
configuration with the credentials redacted; keep it that way.

Captured HTTP responses are fine to commit verbatim — they contain no credentials.

## Building a binary

```bash
uv run pyinstaller pharmparser.spec
```

Entry points live in `packaging/` because PyInstaller runs its entry script as
`__main__`, which would break the package's relative imports. PyInstaller 6.15 is
the floor for Python 3.14.

## Where to read next

[`docs/REFACTOR_PLAN.md`](docs/REFACTOR_PLAN.md) records every bug found in the
original code, the target architecture, and what each phase changed and why.
