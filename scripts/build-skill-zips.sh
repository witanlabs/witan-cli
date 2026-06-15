#!/usr/bin/env bash
# Build one uploadable zip per skill under skills/ for the rolling "skills"
# GitHub release (see .github/workflows/skills-publish.yml).
#
# Layout is folder-at-root — each zip contains <name>/SKILL.md — which is what
# claude.ai's skill uploader requires. It also requires the folder name to
# match the skill's frontmatter name, so that is validated here.
#
# Requires yq (preinstalled on GitHub runners; locally: brew install yq).
#
# Usage: scripts/build-skill-zips.sh <outdir>
set -euo pipefail

outdir="${1:?usage: build-skill-zips.sh <outdir>}"
mkdir -p "$outdir"
outdir="$(cd "$outdir" && pwd)"

cd "$(dirname "$0")/.."

failed=0
for dir in skills/*/; do
  name="$(basename "$dir")"

  if [ ! -f "${dir}SKILL.md" ]; then
    echo "error: ${dir} has no SKILL.md" >&2
    failed=1
    continue
  fi

  fm_name="$(yq --front-matter=extract '.name // ""' "${dir}SKILL.md")"
  if [ "$fm_name" != "$name" ]; then
    echo "error: ${dir} frontmatter name \"${fm_name}\" does not match the folder name — claude.ai's uploader requires them to match" >&2
    failed=1
    continue
  fi

  rm -f "${outdir}/${name}.zip" # zip(1) updates an existing archive in place
  (cd skills && zip -q -r -X "${outdir}/${name}.zip" "$name" --exclude '*.DS_Store')
  echo "built ${name}.zip"
done

if [ "$failed" -ne 0 ]; then
  exit 1
fi
