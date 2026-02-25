#!/bin/sh
set -eu

REPO="${WITAN_REPO:-witanlabs/witan-cli}"
REQUESTED_VERSION="${WITAN_VERSION:-latest}"
INSTALL_DIR="${WITAN_INSTALL_DIR:-/usr/local/bin}"
RETRY_ATTEMPTS="${WITAN_INSTALL_RETRIES:-120}"
RETRY_DELAY_SECONDS="${WITAN_INSTALL_RETRY_DELAY:-5}"

need_cmd() {
  if ! command -v "$1" >/dev/null 2>&1; then
    echo "missing required command: $1" >&2
    exit 1
  fi
}

need_cmd curl
need_cmd awk
need_cmd tar
need_cmd install

resolve_tag() {
  if [ "${REQUESTED_VERSION}" != "latest" ]; then
    case "${REQUESTED_VERSION}" in
      v*) echo "${REQUESTED_VERSION}" ;;
      *) echo "v${REQUESTED_VERSION}" ;;
    esac
    return
  fi

  latest_location="$(
    curl -fsSI "https://github.com/${REPO}/releases/latest" \
      | tr -d '\r' \
      | awk 'tolower($1) == "location:" { print $2 }' \
      | tail -n 1
  )"
  tag="${latest_location##*/}"
  if [ -z "${tag}" ] || [ "${tag}" = "latest" ]; then
    echo "failed to resolve latest release tag for ${REPO}" >&2
    exit 1
  fi
  echo "${tag}"
}

platform_asset() {
  os="$(uname -s | tr '[:upper:]' '[:lower:]')"
  arch="$(uname -m)"

  case "${arch}" in
    x86_64|amd64) arch="amd64" ;;
    arm64|aarch64) arch="arm64" ;;
    *)
      echo "unsupported architecture: ${arch}" >&2
      exit 1
      ;;
  esac

  case "${os}" in
    darwin) echo "witan-darwin-${arch}.tar.gz" ;;
    linux) echo "witan-linux-${arch}.tar.gz" ;;
    *)
      echo "unsupported operating system: ${os}" >&2
      exit 1
      ;;
  esac
}

download_with_retry() {
  url="$1"
  dest="$2"
  attempt=1

  while [ "${attempt}" -le "${RETRY_ATTEMPTS}" ]; do
    if curl -fsSL "${url}" -o "${dest}"; then
      return 0
    fi

    if [ "${attempt}" -eq "${RETRY_ATTEMPTS}" ]; then
      break
    fi

    sleep "${RETRY_DELAY_SECONDS}"
    attempt=$((attempt + 1))
  done

  echo "failed to download ${url} after ${RETRY_ATTEMPTS} attempts" >&2
  return 1
}

sha256_file() {
  file="$1"
  if command -v sha256sum >/dev/null 2>&1; then
    sha256sum "${file}" | awk '{print $1}'
    return
  fi
  if command -v shasum >/dev/null 2>&1; then
    shasum -a 256 "${file}" | awk '{print $1}'
    return
  fi
  echo "missing sha256 tool (need sha256sum or shasum)" >&2
  exit 1
}

TAG="$(resolve_tag)"
ASSET="$(platform_asset)"
DOWNLOAD_BASE="https://github.com/${REPO}/releases/download/${TAG}"

WORK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/witan-install.XXXXXX")"
trap 'rm -rf "${WORK_DIR}"' EXIT INT TERM

CHECKSUMS_FILE="${WORK_DIR}/witan-checksums.txt"
ARCHIVE_FILE="${WORK_DIR}/${ASSET}"

download_with_retry "${DOWNLOAD_BASE}/witan-checksums.txt" "${CHECKSUMS_FILE}"
download_with_retry "${DOWNLOAD_BASE}/${ASSET}" "${ARCHIVE_FILE}"

EXPECTED_SHA="$(
  awk -v asset="${ASSET}" '$2 == asset { print $1 }' "${CHECKSUMS_FILE}" | head -n 1
)"
if [ -z "${EXPECTED_SHA}" ]; then
  echo "checksum entry for ${ASSET} not found in witan-checksums.txt" >&2
  exit 1
fi

ACTUAL_SHA="$(sha256_file "${ARCHIVE_FILE}")"
if [ "${EXPECTED_SHA}" != "${ACTUAL_SHA}" ]; then
  echo "checksum mismatch for ${ASSET}" >&2
  echo "expected: ${EXPECTED_SHA}" >&2
  echo "actual:   ${ACTUAL_SHA}" >&2
  exit 1
fi

tar -xzf "${ARCHIVE_FILE}" -C "${WORK_DIR}"
if [ ! -f "${WORK_DIR}/witan" ]; then
  echo "archive did not contain expected binary: witan" >&2
  exit 1
fi

install_local() {
  target_dir="$1"
  mkdir -p "${target_dir}" && install -m 0755 "${WORK_DIR}/witan" "${target_dir}/witan"
}

TARGET_PATH=""
if install_local "${INSTALL_DIR}" 2>/dev/null; then
  TARGET_PATH="${INSTALL_DIR}/witan"
elif command -v sudo >/dev/null 2>&1; then
  sudo mkdir -p "${INSTALL_DIR}"
  sudo install -m 0755 "${WORK_DIR}/witan" "${INSTALL_DIR}/witan"
  TARGET_PATH="${INSTALL_DIR}/witan"
else
  FALLBACK_DIR="${HOME}/.local/bin"
  if install_local "${FALLBACK_DIR}"; then
    TARGET_PATH="${FALLBACK_DIR}/witan"
    echo "installed to ${TARGET_PATH}" >&2
    echo "add ${FALLBACK_DIR} to PATH if needed" >&2
  else
    echo "failed to install witan to ${INSTALL_DIR} and ${FALLBACK_DIR}" >&2
    exit 1
  fi
fi

echo "installed ${TARGET_PATH} from ${TAG}"
"${TARGET_PATH}" --version
