#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-}"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUTPUT_DIR="${2:-${ROOT_DIR}/dist}"

if [[ -z "${VERSION}" ]]; then
  echo "usage: $0 <version-tag> [output-dir]" >&2
  echo "example: $0 v1.2.3 ./dist" >&2
  exit 1
fi

if ! [[ "${VERSION}" =~ ^v[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
  echo "PyPI wheel publishing currently supports stable tags only (vX.Y.Z). Got: ${VERSION}" >&2
  exit 1
fi

if ! command -v go-to-wheel >/dev/null 2>&1; then
  echo "go-to-wheel is required. Install with: python -m pip install go-to-wheel" >&2
  exit 1
fi

PYPI_VERSION="${VERSION#v}"
mkdir -p "${OUTPUT_DIR}"
# Remove stale wheel artifacts from prior runs when reusing an output directory.
rm -f "${OUTPUT_DIR}"/witan-*.whl

if [[ -z "${GOCACHE:-}" ]]; then
  export GOCACHE="${ROOT_DIR}/.tmp/gocache"
fi
mkdir -p "${GOCACHE}"

go-to-wheel "${ROOT_DIR}" \
  --name witan \
  --entry-point witan \
  --version "${PYPI_VERSION}" \
  --output-dir "${OUTPUT_DIR}" \
  --description "Witan spreadsheet CLI for coding agents" \
  --license "Apache-2.0" \
  --author "Witan Labs" \
  --url "https://github.com/witanlabs/witan-cli" \
  --readme "${ROOT_DIR}/README.md" \
  --set-version-var "github.com/witanlabs/witan-cli/cmd.Version"

echo "Built PyPI wheels in ${OUTPUT_DIR}"
