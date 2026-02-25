#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-dev}"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DIST_DIR="${2:-${ROOT_DIR}/dist}"
if [[ "${DIST_DIR}" != /* ]]; then
  DIST_DIR="${PWD}/${DIST_DIR}"
fi

LDFLAGS="-X github.com/witanlabs/witan-cli/cmd.Version=${VERSION}"

mkdir -p "${DIST_DIR}"
rm -rf "${DIST_DIR:?}/"*

tmp_dir="$(mktemp -d)"
trap 'rm -rf "${tmp_dir}"' EXIT

build_tarball() {
  local goos="$1"
  local goarch="$2"
  local asset_name="$3"
  local stage_dir="${tmp_dir}/${goos}-${goarch}"

  mkdir -p "${stage_dir}"
  GOOS="${goos}" GOARCH="${goarch}" go build -ldflags "${LDFLAGS}" -o "${stage_dir}/witan" "${ROOT_DIR}"
  tar -C "${stage_dir}" -czf "${DIST_DIR}/${asset_name}" witan
}

build_zip() {
  local goos="$1"
  local goarch="$2"
  local asset_name="$3"
  local stage_dir="${tmp_dir}/${goos}-${goarch}"

  mkdir -p "${stage_dir}"
  GOOS="${goos}" GOARCH="${goarch}" go build -ldflags "${LDFLAGS}" -o "${stage_dir}/witan.exe" "${ROOT_DIR}"
  (
    cd "${stage_dir}"
    zip -q "${DIST_DIR}/${asset_name}" witan.exe
  )
}

build_tarball darwin arm64 witan-darwin-arm64.tar.gz
build_tarball darwin amd64 witan-darwin-amd64.tar.gz
build_tarball linux amd64 witan-linux-amd64.tar.gz
build_tarball linux arm64 witan-linux-arm64.tar.gz
build_zip windows amd64 witan-windows-amd64.zip
build_zip windows arm64 witan-windows-arm64.zip

(
  cd "${DIST_DIR}"
  shasum -a 256 \
    witan-darwin-arm64.tar.gz \
    witan-darwin-amd64.tar.gz \
    witan-linux-amd64.tar.gz \
    witan-linux-arm64.tar.gz \
    witan-windows-amd64.zip \
    witan-windows-arm64.zip > witan-checksums.txt
)

echo "Built release artifacts in ${DIST_DIR}"
