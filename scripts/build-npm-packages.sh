#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-}"
ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUTPUT_DIR="${2:-${ROOT_DIR}/npm-dist}"
NODE_DIR="${ROOT_DIR}/node"

if [[ -z "${VERSION}" ]]; then
  echo "usage: $0 <version-tag> [output-dir]" >&2
  echo "example: $0 v1.2.3 ./npm-dist" >&2
  exit 1
fi

NPM_VERSION="${VERSION#v}"
mkdir -p "${OUTPUT_DIR}"
rm -f "${OUTPUT_DIR}"/*.tgz

if [[ -z "${GOCACHE:-}" ]]; then
  export GOCACHE="${ROOT_DIR}/.tmp/gocache"
fi
mkdir -p "${GOCACHE}"

# Build main package in a temp directory to avoid mutating working tree
WORK_DIR=$(mktemp -d)
trap "rm -rf ${WORK_DIR}" EXIT

cp -r "${NODE_DIR}" "${WORK_DIR}/witan"
cd "${WORK_DIR}/witan"

# Remove node_modules to avoid symlink issues in temp directory
rm -rf node_modules

# Update version without modifying the original package.json
# Use --allow-same-version in case the version is already set
npm version "${NPM_VERSION}" --no-git-tag-version --allow-same-version

# Update optionalDependencies to match the new version
# (npm version only updates the top-level version field)
jq --arg v "${NPM_VERSION}" '.optionalDependencies |= with_entries(.value = $v)' package.json > package.json.tmp \
  && mv package.json.tmp package.json

# Fresh install to ensure all dependencies are properly installed
npm install
npm run build
npm pack --pack-destination "${OUTPUT_DIR}"

# Build platform packages
build_platform_package() {
  local goos="$1"
  local goarch="$2"
  local cpu="$3"
  local libc="${4:-}"

  # Determine npm os name
  local npm_os
  case "${goos}" in
    darwin) npm_os="darwin" ;;
    linux) npm_os="linux" ;;
    windows) npm_os="win32" ;;
  esac

  # Build package suffix
  local pkg_suffix="${npm_os}-${cpu}"
  [[ -n "${libc}" ]] && pkg_suffix+="-${libc}"

  # Determine binary extension
  local ext=""
  [[ "${goos}" == "windows" ]] && ext=".exe"

  local pkg_dir="${WORK_DIR}/@witan/${pkg_suffix}"
  mkdir -p "${pkg_dir}/bin"

  echo "Building @witan/${pkg_suffix}..."

  # Build Go binary (run from ROOT_DIR where go.mod lives)
  (cd "${ROOT_DIR}" && GOOS="${goos}" GOARCH="${goarch}" CGO_ENABLED=0 \
    go build \
      -trimpath \
      -ldflags "-s -w -X github.com/witanlabs/witan-cli/cmd.Version=${NPM_VERSION}" \
      -o "${pkg_dir}/bin/witan${ext}" \
      .)

  # Set executable permission
  [[ "${goos}" != "windows" ]] && chmod 0755 "${pkg_dir}/bin/witan"

  # Build description with optional libc suffix
  local desc="Witan CLI binary for ${npm_os} ${cpu}"
  [[ -n "${libc}" ]] && desc+=" (${libc})"

  # Create package.json
  cat > "${pkg_dir}/package.json" <<EOF
{
  "name": "@witan/${pkg_suffix}",
  "version": "${NPM_VERSION}",
  "description": "${desc}",
  "license": "Apache-2.0",
  "repository": {
    "type": "git",
    "url": "https://github.com/witanlabs/witan-cli.git"
  },
  "os": ["${npm_os}"],
  "cpu": ["${cpu}"],
  "bin": {
    "witan": "bin/witan${ext}"
  }
}
EOF

  # Pack
  (cd "${pkg_dir}" && npm pack --pack-destination "${OUTPUT_DIR}")
}

# Build all platform packages
build_platform_package darwin arm64 arm64
build_platform_package darwin amd64 x64
build_platform_package linux amd64 x64
build_platform_package linux arm64 arm64
build_platform_package linux amd64 x64 musl
build_platform_package linux arm64 arm64 musl
build_platform_package windows amd64 x64
build_platform_package windows arm64 arm64

echo "Built npm packages in ${OUTPUT_DIR}:"
ls -la "${OUTPUT_DIR}"/*.tgz
