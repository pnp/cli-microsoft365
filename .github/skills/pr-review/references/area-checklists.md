# Area-Specific Checklists

Apply these in addition to the general checklist when the PR touches the corresponding area.

## SPFx upgrade and doctor rules

Code under `src/m365/spfx/commands/project/` (`project-upgrade/`, `project-doctor/`, rules, and `test-projects/`).

- Reuse existing rules instead of adding near-duplicates. For example, update the Node.js `engines` requirement via `FN021003_PKG_engines_node` rather than a new rule. Search `rules/` for a rule with the same ID prefix or purpose first.
- When a command's support range is intentionally extended, add the new SPFx version to that command's `supportedVersions` and its related docs/tests. Do not require unrelated commands such as `project-externalize` to match the upgrade/doctor support matrix when their implementations have a separately validated range.
- The version-specific upgrade/doctor file (e.g. `upgrade-<version>.ts`, `doctor-<version>.ts`) lists rules in the expected order and includes every rule the SPFx release requires.
- Rules that modify `package.json` overrides, resolutions, or arrays (e.g. `excludeDirectories`) don't overwrite or drop user-defined entries, and don't remove a value without re-adding the required one.
- Matching `test-projects/` fixtures are added or updated, and specs exercise both the "fix needed" and "already compliant" cases.
- Rule severity, resolution text, and file paths match the conventions of existing rules.

## Docs site

Changes under `docs/` other than command reference pages (Docusaurus config, plugins, components, theme, build scripts).

- `docs/docusaurus.config.ts` changes keep existing behavior (presets, plugins, redirects, search, versioning) unless the change is intended; fallback/default values match the previous defaults.
- Every package imported by docs code is declared as a direct dependency in `docs/package.json` (don't rely on transitive dependencies such as `@docusaurus/utils`).
- Build performance: no unbounded concurrency (e.g. spawning a `git` process per file without a limit), no repeated expensive work per page; consider the impact on `npm run build` time in `docs/`.
- Derived metadata is correct (e.g. a page's creation date comes from the first commit, not the last update).
- The docs build (`cd docs && npm ci && npm run build`) succeeds without new warnings or broken links.

## Telemetry and infrastructure

Changes to `src/telemetry.ts`, `src/googleAnalytics.ts`, `src/request.ts`, CLI startup, CI workflows, Docker images, or other cross-cutting infrastructure.

- Privacy: any change to what telemetry collects or where it's sent is documented in `docs/docs/about/telemetry.mdx`. No personal data, tenant data, or option values that may contain them are sent.
- Secrets: no API secrets, keys, or tokens committed in source; anything shipped in the public package is treated as public.
- Network calls have timeouts and fail silently (telemetry must never break or slow down a command).
- External API limits and types are respected (e.g. Google Analytics 4 Measurement Protocol: parameter name length limits, maximum number of parameters per event, supported value types — booleans aren't supported and must be converted). Custom parameters don't overwrite reserved or existing metadata (e.g. version, shell).
- Existing environment detection and tests (e.g. Docker, CI, shell detection) aren't removed without a replacement.
- CI/workflow changes keep required checks and pinned versions intact.
