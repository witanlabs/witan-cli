# witan-cli

Command-line client for Witan APIs.

## Prerequisites

- Go (version from `go.mod`)

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
```

The local binary is written to `./witan`.

## Release Process

Releases are now handled directly in this repository by GitHub Actions:

- Workflow: `.github/workflows/witan-cli-release.yml`
- Trigger tags: `v*` (for example `v1.2.3`)
- Artifacts:
  - `witan-darwin-arm64.tar.gz`
  - `witan-darwin-amd64.tar.gz`
  - `witan-linux-amd64.tar.gz`
  - `witan-linux-arm64.tar.gz`
  - `witan-windows-amd64.zip`
  - `witan-windows-arm64.zip`
  - `witan-checksums.txt`

Example release:

```bash
git tag v1.2.3
git push origin v1.2.3
```

You can also run the release workflow manually with `workflow_dispatch` and provide `version` (for example `v1.2.3`).

## CI

Go CI runs in `.github/workflows/golang.yml` on pushes to `main` and pull requests.
