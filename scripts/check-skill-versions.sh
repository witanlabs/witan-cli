#!/usr/bin/env bash
# CI gate: any change under skills/<name>/ must bump metadata.version in that
# skill's SKILL.md. See skills/README.md.
#
# Requires yq (preinstalled on GitHub runners; locally: brew install yq).
#
# Usage: scripts/check-skill-versions.sh <base-ref>
#   e.g. scripts/check-skill-versions.sh origin/main
set -euo pipefail

base_ref="${1:?usage: check-skill-versions.sh <base-ref>}"
base="$(git merge-base "$base_ref" HEAD)"
failed=0

frontmatter() { # <rev> <dir> <yq-expr>
  git show "$1:$2/SKILL.md" 2>/dev/null | yq --front-matter=extract "$3" - 2>/dev/null || true
}

is_plain_version() { # e.g. "1.0.0" — no prerelease or build suffixes
  [[ "$1" =~ ^[0-9]+(\.[0-9]+)*$ ]]
}

for dir in $(git diff --name-only "$base" HEAD -- skills/ | cut -d/ -f1-2 | sort -u); do
  [ -f "$dir/SKILL.md" ] || continue # skill deleted in this PR

  new="$(frontmatter HEAD "$dir" '.metadata.version // ""')"
  old="$(frontmatter "$base" "$dir" '.metadata.version // ""')"

  if [ "$(frontmatter HEAD "$dir" 'has("version")')" = "true" ]; then
    echo "error: $dir SKILL.md has a top-level version key — the skills spec rejects it; use metadata.version" >&2
    failed=1
  fi

  if [ -z "$new" ]; then
    echo "error: $dir changed but SKILL.md has no metadata.version" >&2
    failed=1
  elif ! is_plain_version "$new"; then
    echo "error: $dir metadata.version \"$new\" is not a plain dotted version like \"1.0.0\"" >&2
    failed=1
  elif [ -z "$old" ] || ! is_plain_version "$old"; then
    echo "$dir: new at $new"
  elif [ "$new" = "$old" ] || [ "$(printf '%s\n%s\n' "$old" "$new" | sort -V | tail -n 1)" != "$new" ]; then
    echo "error: $dir changed but metadata.version did not increase ($old -> $new) — bump it" >&2
    failed=1
  else
    echo "$dir: $old -> $new"
  fi
done

exit "$failed"
