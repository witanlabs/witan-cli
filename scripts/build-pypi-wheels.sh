#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-}"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUTPUT_DIR="${2:-${ROOT_DIR}/dist}"
BINARY_DIR="${ROOT_DIR}/python/witan/bin"

if [[ -z "${VERSION}" ]]; then
  echo "usage: $0 <version-tag> [output-dir]" >&2
  echo "example: $0 v1.2.3 ./dist" >&2
  exit 1
fi

if ! [[ "${VERSION}" =~ ^v[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
  echo "PyPI wheel publishing currently supports stable tags only (vX.Y.Z). Got: ${VERSION}" >&2
  exit 1
fi

PYPI_VERSION="${VERSION#v}"
mkdir -p "${OUTPUT_DIR}" "${BINARY_DIR}"
rm -f "${OUTPUT_DIR}"/witan-*.whl

if [[ -z "${GOCACHE:-}" ]]; then
  export GOCACHE="${ROOT_DIR}/.tmp/gocache"
fi
mkdir -p "${GOCACHE}"

cleanup_staged_binary() {
  rm -f "${BINARY_DIR}/witan" "${BINARY_DIR}/witan.exe"
}
trap cleanup_staged_binary EXIT

build_wheel() {
  local goos="$1"
  local goarch="$2"
  local platform_tag="$3"
  local extension="${4:-}"
  local binary_path="${BINARY_DIR}/witan${extension}"

  cleanup_staged_binary
  echo "Building ${platform_tag} wheel"

  GOOS="${goos}" GOARCH="${goarch}" CGO_ENABLED=0 \
    go build \
      -trimpath \
      -ldflags "-s -w -X github.com/witanlabs/witan-cli/cmd.Version=${PYPI_VERSION}" \
      -o "${binary_path}" \
      "${ROOT_DIR}"

  if [[ "${goos}" != "windows" ]]; then
    chmod 0755 "${binary_path}"
  fi

  rm -rf "${ROOT_DIR}/build" "${ROOT_DIR}/python/witan.egg-info" "${ROOT_DIR}/witan.egg-info"
  WITAN_PY_VERSION="${PYPI_VERSION}" \
    WITAN_WHEEL_PLAT_NAME="${platform_tag}" \
    python -m build --wheel --no-isolation --outdir "${OUTPUT_DIR}" "${ROOT_DIR}"
}

build_wheel darwin arm64 macosx_11_0_arm64
build_wheel darwin amd64 macosx_10_15_x86_64
build_wheel linux amd64 manylinux_2_17_x86_64
build_wheel linux arm64 manylinux_2_17_aarch64
build_wheel linux amd64 musllinux_1_2_x86_64
build_wheel linux arm64 musllinux_1_2_aarch64
build_wheel windows amd64 win_amd64 .exe
build_wheel windows arm64 win_arm64 .exe

cleanup_staged_binary
rm -rf "${ROOT_DIR}/build" "${ROOT_DIR}/python/witan.egg-info" "${ROOT_DIR}/witan.egg-info"

echo "Built PyPI wheels in ${OUTPUT_DIR}"
