---
name: patch-vulnerabilities
description: >-
  This skill should be used when the user asks to "patch vulnerabilities",
  "fix npm audit issues", "update vulnerable dependencies", "scan and fix
  vulnerabilities", "run npm audit and patch", "fix security vulnerabilities",
  "update outdated dependencies", "audit deps", or needs to iteratively scan,
  patch, and re-verify npm dependency vulnerabilities and outdated packages
  with a cooldown safety check.
---

# Patch npm Dependencies

Comprehensively scan npm dependencies for vulnerabilities, outdated packages,
and stale overrides. Patch eligible ones (respecting a 7-day publish-age
cooldown), verify fixes, and repeat until no new patchable issues remain.

## Prerequisites

- Node.js and npm installed
- A project with `package.json` and `npm-shrinkwrap.json` or `package-lock.json`
- Network access to the npm registry

## Workflow

Execute the phases below in a loop. Each pass through the loop is one
**patch cycle**. Continue cycling until the termination condition is met.

---

### Phase 1 — Scan (three checks)

Run all three scans to get full coverage of dependency issues.

#### 1A — Security vulnerabilities (`npm audit`)

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

#### 1B — Outdated direct dependencies (`npm outdated`)

```bash
npm outdated --json 2>/dev/null
```

Parse the JSON output. For each entry, extract:
- Package name
- Current version
- Wanted version (max satisfying the declared range)
- Latest version (latest on registry)
- Whether upgrading to latest requires a major version bump

**Note:** `npm outdated` only reports direct dependencies listed in
`dependencies` and `devDependencies`. It does NOT report outdated overrides.

#### 1C — Outdated overrides (manual registry check)

Read the `overrides` field from `package.json` (including nested overrides).
For each override entry:

```bash
npm view <package> version --json
```

Compare the pinned override version against the latest version on the
registry. If the latest is newer, the override is outdated.

Also check nested overrides. Overrides can be structured as:
```json
{
  "overrides": {
    "parent-package": {
      "child-package": "1.2.3"
    }
  }
}
```
In this case, check `child-package` against the registry.

#### Phase 1 summary

If no issues are found across all three checks, stop — the project is clean.

Otherwise, present findings in tables:

**Security vulnerabilities:**

| Package | Severity | Current | Fix available | Type | Top-level ancestor |
|---------|----------|---------|---------------|------|--------------------|

**Outdated direct dependencies:**

| Package | Current | Wanted | Latest | Major bump? |
|---------|---------|--------|--------|-------------|

**Outdated overrides:**

| Package | Override version | Latest | Context (parent) |
|---------|-----------------|--------|------------------|

---

### Phase 2 — Check cooldown eligibility

For every package that has an available upgrade (from any of the three
checks), verify the target version's publish date against the **7-day
cooldown rule**: the target version must have been published at least 7 days
ago.

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
**If `age_days < 7`**: skip this package for now and report it as "cooling
down" with the date it becomes eligible.

Report cooldown status to the user:

| Package | Target version | Published | Eligible | Eligible date |
|---------|----------------|-----------|----------|---------------|

If no packages are eligible, stop — all remaining issues are in cooldown.
Report the earliest eligibility date.

---

### Phase 3 — Patch

Apply fixes for all eligible packages, grouped by source.

#### 3A — Security vulnerabilities

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

#### 3B — Outdated direct dependencies

For each outdated direct dependency that passed cooldown:

1. **Patch/minor update** (wanted version matches latest, no major bump):
   ```bash
   npm install <package>@<latest-version>
   ```

2. **Major version update** (latest requires a major bump):
   ```bash
   npm install <package>@<latest-version>
   ```
   Flag this to the user as a **breaking change** and note it requires
   additional testing.

#### 3C — Outdated overrides

For each outdated override that passed cooldown:

1. Update the version in the `overrides` section of `package.json` to the
   latest version.

2. Run `npm install` to apply the override resolution.

3. Flag the change to the user — overrides bypass the parent package's
   declared compatibility range and may cause runtime issues.

4. If the override was originally added to fix a vulnerability, check whether
   the parent package now ships a version that includes the fix natively.
   If so, recommend removing the override entirely and updating the parent
   package instead.

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
- Commit message format:
  - For vulnerability fixes: `fix: upgrade <package> to <version> to fix <severity> vulnerability`
  - For outdated dependencies: `fix: upgrade <package> from <old> to <new>`
  - For outdated overrides: `fix: upgrade <package> override from <old> to <new>`

---

### Phase 4 — Re-scan and loop

Return to **Phase 1**. Run all three checks again to verify fixes and detect
any newly-revealed issues.

---

### Termination conditions

Stop the loop when any of these is true:

1. All three scans report **zero issues**
2. All remaining issues have target versions **in cooldown** (< 7 days old)
3. All remaining issues have **no fix available**
4. A patch cycle produced **zero successful patches** (nothing new was fixed)

---

### Final report

After the loop ends, present a summary:

```
## Dependency Patch Summary

Patch cycles completed: N

### Packages patched
- Security fixes:
  - Direct: list with versions
  - Transitive (via ancestor update): list
  - Transitive (via override): list
- Outdated direct dependencies updated: list with old → new versions
- Outdated overrides updated: list with old → new versions

### Remaining issues
- Vulnerabilities in cooldown (eligible on <date>): list
- Outdated packages in cooldown (eligible on <date>): list
- No fix available: list
- Transitive, waiting on upstream: list with top-level ancestor

### Overrides
- Added: list
- Updated: list
- Recommend removing (upstream fix available): list
(Review overrides periodically and remove when upstream fixes land)
```

---

## Important notes

- Always run tests between patches to catch breakage early.
- Commit each patch individually for easy rollback.
- When a fix requires a major version bump, warn the user explicitly.
- Never force-install a version published less than 7 days ago.
- If `npm-shrinkwrap.json` exists, include it in commits alongside
  `package-lock.json`.
- `npm audit` only catches security vulnerabilities — not deprecated or
  outdated packages.
- `npm outdated` only catches outdated direct dependencies — not overrides.
- Overrides are invisible to both `npm audit` and `npm outdated` — they must
  be checked manually against the registry.
