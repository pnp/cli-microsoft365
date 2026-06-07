---
name: patch-vulnerabilities
description: >-
  This skill should be used when the user asks to "patch vulnerabilities",
  "fix npm audit issues", "update vulnerable dependencies", "update outdated
  dependencies", "scan and fix vulnerabilities", "run npm audit and patch",
  "fix security vulnerabilities", "update npm packages", "check for outdated
  packages", or needs to comprehensively scan, patch, and re-verify npm
  dependencies with a cooldown safety check. Covers three blind spots:
  npm audit (security), npm outdated (staleness), and override checks.
---

# Patch npm Dependencies

Comprehensively scan npm dependencies for security vulnerabilities, outdated
packages, and stale overrides. Patch eligible ones (respecting a 7-day
publish-age cooldown), verify fixes, and repeat until no new patchable
issues remain.

Each tool catches a different blind spot:
1. `npm audit` — known **security vulnerabilities** only. Does not flag
   deprecated or outdated packages.
2. `npm outdated` — catches outdated **direct dependencies** only. Ignores
   overrides and does not flag vulnerabilities.
3. **Override checks** — overrides are invisible to both tools. You must
   manually check each override version against the npm registry.

All three checks are required for full coverage.

## Prerequisites

- Node.js and npm installed
- A project with `package.json` and `npm-shrinkwrap.json` or `package-lock.json`
- Network access to the npm registry

## Workflow

Execute the phases below in a loop. Each pass through the loop is one
**patch cycle**. Continue cycling until the termination condition is met.

---

### Phase 1 — Scan

Run all three scans and combine results into a single unified list of
actionable items.

#### 1A. Security vulnerabilities (`npm audit`)

```bash
npm audit --json 2>/dev/null
```

Parse the JSON output. Extract from the `vulnerabilities` object:
- Package name
- Current version (from `nodes` or `range`)
- Severity (`critical`, `high`, `moderate`, `low`)
- Fix available (`via` and `fixAvailable` fields)
- Dependency type: **direct** or **transitive**

To classify each vulnerability:
- **Direct**: the package appears in `dependencies` or `devDependencies` in
  `package.json`
- **Transitive**: the package does NOT appear in `package.json` — it is pulled
  in by a direct dependency. The `npm audit --json` output shows the dependency
  chain in the `via` and `effects` fields. Identify the **top-level ancestor**
  (the direct dependency that pulls in the vulnerable transitive package).

#### 1B. Outdated direct dependencies (`npm outdated`)

```bash
npm outdated --json 2>/dev/null
```

Parse the JSON output. For each entry extract:
- Package name
- `current` version (what is installed)
- `wanted` version (latest that satisfies the semver range in `package.json`)
- `latest` version (latest on the registry)

An entry is actionable when `current` differs from `wanted` (semver-compatible
update available). Note entries where `wanted` differs from `latest` — these
indicate a major version is available but not semver-compatible with the
declared range; flag these separately.

#### 1C. Stale overrides

Read the `overrides` object from `package.json`. Flatten all nested overrides
to a list of `{ parent, package, pinnedVersion }` tuples. For each, query the
npm registry for the latest version:

```bash
npm view <package> version
```

Compare the pinned version to the latest. An override is stale when
`pinnedVersion !== latest`. Classify the gap:
- **Patch/minor**: likely safe to update
- **Major**: flag as a **breaking change**

#### Combined scan summary

Present **all** findings to the user in a single table:

| Package | Source | Current | Target | Type | Severity/Notes |
|---------|--------|---------|--------|------|----------------|

Where **Source** is one of: `audit`, `outdated`, `override`.

If all three scans find zero actionable items, stop — the project is clean.

---

### Phase 2 — Check cooldown eligibility

For every actionable item from Phase 1 that has an update target, verify the
target version's publish date against the **7-day cooldown rule**: the target
version must have been published at least 7 days ago.

Query the npm registry for each package:

```bash
npm view <package> time --json
```

This returns a JSON object mapping version strings to ISO 8601 timestamps.
Find the entry for the target version. Calculate the age:

```
age_days = (now - publish_date) / 86400
```

**If `age_days >= 7`**: the package is eligible for patching.
**If `age_days < 7`**: skip this package for now and report it as "cooling down"
with the date it becomes eligible.

Report cooldown status to the user:

| Package | Source | Target version | Published | Age | Eligible | Eligible date |
|---------|--------|---------------|-----------|-----|----------|---------------|

If no packages are eligible, stop — all remaining items are in cooldown or
have no fix. Report the earliest eligibility date.

---

### Phase 3 — Patch

For each eligible package, apply the fix. The strategy depends on the source
and dependency type.

#### 3A. Security vulnerabilities (from `npm audit`)

##### Direct dependencies

1. **Direct fix** — If `npm audit fix` can resolve it without breaking changes:
   ```bash
   npm audit fix --dry-run 2>/dev/null
   ```
   Review the dry-run output. If changes look safe, apply:
   ```bash
   npm audit fix
   ```

2. **Targeted update** — If `npm audit fix` cannot resolve it or introduces
   breaking changes, update the specific package:
   ```bash
   npm install <package>@<fix-version>
   ```

3. **Major version update** — If the fix requires a major version bump:
   ```bash
   npm install <package>@<fix-version>
   ```
   Flag this to the user as a **breaking change** and note it requires
   additional testing.

##### Transitive dependencies

Transitive vulnerabilities cannot be fixed by installing the vulnerable
package directly. Work through the dependency chain instead:

1. **Update the top-level ancestor** — Often, updating the direct dependency
   that pulls in the vulnerable transitive package resolves the issue:
   ```bash
   npm install <top-level-ancestor>@latest
   ```
   Apply the cooldown check to this version too.

2. **npm audit fix** — This resolves the dependency tree automatically when
   a compatible version exists:
   ```bash
   npm audit fix
   ```

3. **npm overrides** — If the top-level ancestor has not released a fix yet,
   use an npm override to force the transitive dependency to the fixed version.
   Add to `package.json`:
   ```json
   {
     "overrides": {
       "<vulnerable-package>": "<fix-version>"
     }
   }
   ```
   Then run `npm install` to apply. Apply the cooldown check to the override
   version. Flag overrides to the user — they bypass the parent package's
   declared compatibility range and may cause runtime issues.

4. **No fix available** — If the top-level ancestor pins the vulnerable
   version and no override is safe, report the vulnerability as unfixable
   for now. Note the top-level ancestor and suggest the user open an issue
   or PR upstream.

#### 3B. Outdated direct dependencies (from `npm outdated`)

For each outdated package where `current != wanted`:

```bash
npm install <package>@<wanted-version>
```

This is a semver-compatible update within the declared range in `package.json`,
so it should be safe. If `wanted != latest` and the user wants to pursue the
latest (major) version, flag it as a **breaking change**.

#### 3C. Stale overrides

For each stale override, update the pinned version in `package.json`'s
`overrides` section, then run:

```bash
npm install
```

- **Patch/minor updates**: update the version string directly.
- **Major updates**: flag to the user as a **breaking change**. Since overrides
  bypass the parent package's declared compatibility range, a major bump is
  especially risky. Only apply if the user confirms, or skip and report.

#### After each patch

Build the project and run the test suite to verify nothing broke:

```bash
npm run build && npm test
```

If the build or tests fail after a patch:
- Revert the change: `git checkout -- package.json npm-shrinkwrap.json package-lock.json && npm install`
- Report the failure to the user with the test output
- Continue to the next eligible package

If tests pass, commit the change:
- Stage: `git add package.json npm-shrinkwrap.json package-lock.json`
- Commit with a descriptive message, e.g.:
  - Audit fix: `fix: upgrade <package> to <version> to fix <severity> vulnerability`
  - Outdated dep: `fix: upgrade <package> from <old> to <new>`
  - Override update: `fix: update <package> override from <old> to <new>`

---

### Phase 4 — Re-scan and loop

Return to **Phase 1**. Run all three scans again to check for remaining
issues. A previous patch may have resolved transitive issues or introduced
new ones.

---

### Termination conditions

Stop the loop when any of these is true:

1. All three scans (`npm audit`, `npm outdated`, override check) report
   **zero actionable items**
2. All remaining items have targets **in cooldown** (< 7 days old)
3. All remaining items have **no fix available**
4. A patch cycle produced **zero successful patches** (nothing new was fixed)

---

### Final report

After the loop ends, present a summary:

```
## Dependency Patch Summary

Patch cycles completed: N

Packages patched: count
  From npm audit:
    - Direct: list with versions
    - Transitive (via ancestor update): list
    - Transitive (via override): list
  From npm outdated:
    - list with old → new versions
  From override check:
    - list with old → new versions and parent package

Remaining issues: count
  - In cooldown (eligible on <date>): list
  - No fix available: list
  - Major version available (not auto-applied): list
  - Transitive, waiting on upstream: list with top-level ancestor

Overrides: list all current overrides
  - Review periodically and remove when upstream fixes land
```

---

## Important notes

- Always run tests between patches to catch breakage early.
- Commit each patch individually for easy rollback.
- When a fix requires a major version bump, warn the user explicitly.
  Do not auto-apply major version bumps for overrides — flag and skip.
- Never force-install a version published less than 7 days ago.
- If `npm-shrinkwrap.json` exists, include it in commits alongside
  `package-lock.json`.
- Overrides bypass the parent package's declared compatibility range.
  Always flag new or updated overrides to the user.
