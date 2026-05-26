---
name: patch-vulnerabilities
description: >-
  This skill should be used when the user asks to "patch vulnerabilities",
  "fix npm audit issues", "update vulnerable dependencies", "scan and fix
  vulnerabilities", "run npm audit and patch", "fix security vulnerabilities",
  or needs to iteratively scan, patch, and re-verify npm dependency
  vulnerabilities with a cooldown safety check.
---

# Patch npm Vulnerabilities

Iteratively scan npm dependencies for vulnerabilities, patch eligible ones
(respecting a 7-day publish-age cooldown), verify fixes, and repeat until no
new patchable vulnerabilities remain.

## Prerequisites

- Node.js and npm installed
- A project with `package.json` and `npm-shrinkwrap.json` or `package-lock.json`
- Network access to the npm registry

## Workflow

Execute the phases below in a loop. Each pass through the loop is one
**patch cycle**. Continue cycling until the termination condition is met.

### Phase 1 — Scan

Run `npm audit` to identify vulnerabilities:

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

If no vulnerabilities are found, stop — the project is clean.

Summarize findings to the user in a table:

| Package | Severity | Current | Fix available | Type | Top-level ancestor |
|---------|----------|---------|---------------|------|--------------------|

### Phase 2 — Check cooldown eligibility

For each vulnerability that has a fix available, verify the target version's
publish date against the **7-day cooldown rule**: the fix version must have
been published at least 7 days ago.

Query the npm registry for each package:

```bash
npm view <package> time --json
```

This returns a JSON object mapping version strings to ISO 8601 timestamps.
Find the entry for the target fix version. Calculate the age:

```
age_days = (now - publish_date) / 86400
```

**If `age_days >= 7`**: the package is eligible for patching.
**If `age_days < 7`**: skip this package for now and report it as "cooling down"
with the date it becomes eligible.

Report cooldown status to the user:

| Package | Fix version | Published | Eligible | Eligible date |
|---------|-------------|-----------|----------|---------------|

If no packages are eligible, stop — remaining vulnerabilities are all in
cooldown. Report the earliest eligibility date.

### Phase 3 — Patch

For each eligible package, apply the fix. The strategy depends on whether
the vulnerability is in a direct or transitive dependency.

#### Direct dependencies

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

#### Transitive dependencies

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

After patching, build the project and run the test suite to verify nothing broke:

```bash
npm run build && npm test
```

If the build or tests fail after a patch:
- Revert the change: `git checkout -- package.json npm-shrinkwrap.json package-lock.json && npm install`
- Report the failure to the user with the test output
- Continue to the next eligible package

If tests pass, commit the change:
- Stage: `git add package.json npm-shrinkwrap.json package-lock.json`
- Commit with message: `fix: upgrade <package> to <version> to fix <severity> vulnerability`

### Phase 4 — Re-scan and loop

Return to **Phase 1**. Run `npm audit --json` again to check for remaining
vulnerabilities.

### Termination conditions

Stop the loop when any of these is true:

1. `npm audit` reports **zero vulnerabilities**
2. All remaining vulnerabilities have fixes **in cooldown** (< 7 days old)
3. All remaining vulnerabilities have **no fix available**
4. A patch cycle produced **zero successful patches** (nothing new was fixed)

### Final report

After the loop ends, present a summary:

```
## Vulnerability Patch Summary

Patch cycles completed: N
Packages patched: list with versions
  - Direct: list
  - Transitive (via ancestor update): list
  - Transitive (via override): list
Remaining vulnerabilities: count
  - In cooldown (eligible on <date>): list
  - No fix available: list
  - Transitive, waiting on upstream: list with top-level ancestor
Overrides added: list (review periodically and remove when upstream fixes land)
```

## Important notes

- Always run tests between patches to catch breakage early.
- Commit each patch individually for easy rollback.
- When a fix requires a major version bump, warn the user explicitly.
- Never force-install a version published less than 7 days ago.
- If `npm-shrinkwrap.json` exists, include it in commits alongside
  `package-lock.json`.
