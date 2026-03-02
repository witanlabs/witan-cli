#!/usr/bin/env bash
set -euo pipefail

VERSION="${1:-}"
CHANGELOG_FILE="${2:-CHANGELOG.md}"

if [[ -z "${VERSION}" ]]; then
  echo "usage: $0 <version-tag-or-number> [changelog-file]" >&2
  exit 1
fi

if [[ ! -f "${CHANGELOG_FILE}" ]]; then
  echo "changelog file not found: ${CHANGELOG_FILE}" >&2
  exit 1
fi

VERSION="${VERSION#v}"

if grep -Fqx "## ${VERSION}" "${CHANGELOG_FILE}"; then
  echo "changelog already has ${VERSION}; nothing to do"
  exit 0
fi

unreleased_line="$(
  grep -n "^## Unreleased$" "${CHANGELOG_FILE}" | head -n 1 | cut -d: -f1 || true
)"
if [[ -z "${unreleased_line}" ]]; then
  echo "missing '## Unreleased' section in ${CHANGELOG_FILE}" >&2
  exit 1
fi

next_heading_line="$(
  awk -v start="${unreleased_line}" 'NR > start && /^## / { print NR; exit }' "${CHANGELOG_FILE}"
)"

if [[ -n "${next_heading_line}" ]]; then
  unreleased_block="$(sed -n "$((unreleased_line + 1)),$((next_heading_line - 1))p" "${CHANGELOG_FILE}")"
else
  unreleased_block="$(sed -n "$((unreleased_line + 1)),\$p" "${CHANGELOG_FILE}")"
fi

if [[ -z "$(printf '%s\n' "${unreleased_block}" | sed '/^[[:space:]]*$/d')" ]]; then
  echo "unreleased section is empty; nothing to roll"
  exit 0
fi

tmp_file="$(mktemp)"
trap 'rm -f "${tmp_file}"' EXIT

sed -n "1,${unreleased_line}p" "${CHANGELOG_FILE}" > "${tmp_file}"
printf '\n' >> "${tmp_file}"
printf '## %s\n\n' "${VERSION}" >> "${tmp_file}"
printf '%s\n' "${unreleased_block}" >> "${tmp_file}"
printf '\n' >> "${tmp_file}"

if [[ -n "${next_heading_line}" ]]; then
  sed -n "${next_heading_line},\$p" "${CHANGELOG_FILE}" >> "${tmp_file}"
fi

mv "${tmp_file}" "${CHANGELOG_FILE}"
trap - EXIT
echo "rolled unreleased entries into ${VERSION}"
