# witan-cli

Command-line client for Witan spreadsheet APIs.

## Install

### Quick Install Script

```bash
curl -fsSL https://witanlabs.com/agents/install.sh | sh
```

### From GitHub Releases

Download the latest artifacts from:

- https://github.com/witanlabs/witan-cli/releases/latest

Example (macOS Apple Silicon):

```bash
curl -fsSL https://github.com/witanlabs/witan-cli/releases/latest/download/witan-darwin-arm64.tar.gz | tar -xz
install -m 0755 witan /usr/local/bin/witan
```

### From PyPI

Install from PyPI (recommended for sandboxed agent environments):

```bash
# one-shot run without permanent install
uvx witan --help

# persistent install
pip install witan
```

### From Source

Requires Go (version from `go.mod`):

```bash
go install github.com/witanlabs/witan-cli@latest
```

## Quick Start

```bash
# Authenticate (recommended)
witan auth login

# Render a range
witan xlsx render report.xlsx -r "Sheet1!A1:F20"

# Recalculate formulas
witan xlsx calc report.xlsx

# Lint formulas
witan xlsx lint report.xlsx

# Run JS against workbook
witan xlsx exec report.xlsx --expr 'wb.sheet("Summary").cell("A1").value'
```

## What This CLI Covers

`witan-cli` currently exposes four spreadsheet commands:

- `witan xlsx calc`
- `witan xlsx exec`
- `witan xlsx lint`
- `witan xlsx render`

The lower-level Witan spreadsheet runtime supports broader workbook operations; this CLI focuses on the four agent-facing workflows above.

## Auth, Config, and Modes

Authentication can be done via `witan auth login`, `--api-key`, or `WITAN_API_KEY`.

Environment variables:

- `WITAN_API_KEY`: API key (optional when using `witan auth login`)
- `WITAN_API_URL`: API base URL override (default: `https://api.witanlabs.com`)
- `WITAN_STATELESS`: set `1` or `true` to force stateless mode
- `WITAN_CONFIG_DIR`: override config directory (default: `~/.config/witan`)
- `WITAN_MANAGEMENT_API_URL`: management API override for auth login/token exchange

Modes:

- Stateful (default when authenticated): uploads workbook revisions and reuses them across commands
- Stateless (`--stateless` or `WITAN_STATELESS=1`): sends workbook bytes on every request, no server-side file reuse

Limits:

- Workbook inputs must be `<= 25MB`.

## Development

```bash
# build local binary
make build

# run test suite
make test

# static checks
make vet
make format-check

# build release artifacts into dist/
make dist VERSION=v0.1.0

# build PyPI wheels (stable tags only)
make pypi-wheels VERSION=v0.1.0
```

The local binary is written to `./witan`.

## Release Process

Releases are handled by GitHub Actions:

- Publish workflow: `.github/workflows/witan-cli-release.yml` (triggered by pushing `v*` tags)
- Artifacts:
  - `witan-darwin-arm64.tar.gz`
  - `witan-darwin-amd64.tar.gz`
  - `witan-linux-amd64.tar.gz`
  - `witan-linux-arm64.tar.gz`
  - `witan-windows-amd64.zip`
  - `witan-windows-arm64.zip`
  - `witan-install.sh`
  - `witan-*.whl` (PyPI wheels for supported platforms; stable tags only)
  - `witan-checksums.txt`

PyPI publishing:

- Stable tags (`vX.Y.Z`) publish wheels to PyPI using GitHub OIDC trusted publishing.
- Pre-release tags (for example `v1.2.3-rc.1`) skip PyPI publish.

GitHub release publishing:

- The workflow uploads artifacts directly to the matching GitHub Release tag.
- If the release already exists (for example, created in the GitHub UI), assets are attached with `--clobber`.

Cutting a release (UI-driven):

1. Create a GitHub Release in the UI with a new tag `vX.Y.Z` (or prerelease tag `vX.Y.Z-suffix`).
2. Tag push triggers `Witan CLI Release`.
3. The workflow builds artifacts, attaches them to the GitHub Release, and publishes to PyPI for stable tags.
4. For stable tags, verify `witan==X.Y.Z` on PyPI and `witan --version`.

Manual `git tag ... && git push ...` is equivalent to UI tag creation and triggers the same workflow.

## CI

Go CI runs in `.github/workflows/golang.yml` on pushes to `main` and pull requests.
