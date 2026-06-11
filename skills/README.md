# Skills

Agent skills, installed by end users with `npx skills add witanlabs/witan-cli` (and by other clients copying these folders). Installs pull from `main` HEAD — **merging to `main` publishes**, independently of tagged CLI releases.

## Versioning

Each skill declares its own semver in SKILL.md frontmatter:

```yaml
metadata:
  version: "1.2.0"
```

When a PR changes any file under a skill's folder (SKILL.md or `references/`), in that same PR:

1. Bump `metadata.version` — patch: typo/correction; minor: content change; major: restructure or workflow change.
2. Add a `CHANGELOG.md` entry under `## Unreleased`: `- Updated: [Skill] \`<name>\` <version> — <what changed>`.

The `skill-version` CI job fails the PR if a skill changed without a version increase. Skills version independently — bump only what changed. Run the check locally:

```sh
node scripts/check-skill-versions.mjs origin/main
```

## Support

To find out which version a user has installed, have them ask their agent to read the skill's frontmatter. To upgrade: `npx skills update`, or re-run the install command.
