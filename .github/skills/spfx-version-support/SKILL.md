---
name: spfx-version-support
description: 'Add support for a new SPFx version to the CLI. Use when: "add support for SPFx X.Y.Z", "support new SPFx version", "SPFx version upgrade support", "implement SPFx X.Y.Z".'
---

# Add Support for a New SPFx Version

Add support for a new SharePoint Framework (SPFx) version to the three affected commands: `spfx project upgrade`, `spfx project doctor`, and `spfx doctor`. The workflow produces new rule files, updates supported version lists, adds e2e tests, and opens a PR.

## Prerequisites

Before starting, confirm the following from the user:
1. The target SPFx version (e.g. `1.23.2`)
2. The Microsoft release notes URL for that version
3. Whether the user has the SPFx Yeoman generator for that version installed (needed for Phase 3)

If any are missing, **STOP — ask the user** before proceeding.

## Workflow

Execute each phase in order. Do not skip phases.

---

### Phase 1: Research

1. Read the release notes at the URL provided by the user.
2. Determine the **upgrade path**: identify the previous supported version. If any intermediate version was delisted/skipped, note this — the upgrade rules file is named after the target version and upgrades *from* the previous supported version.
3. Identify what changed:
   - **Patch release** (X.Y.Z with Z > 0): typically only `@microsoft/sp-*` package version bumps. No external dep changes (same node, yo, heft, typescript ranges as the prior minor).
   - **Minor/major release**: may introduce new external deps, tooling version bumps, new ESLint rules, SCSS changes, etc.
4. Read the prior minor's upgrade file to understand the baseline rule set:
   ```
   src/m365/spfx/commands/project/project-upgrade/upgrade-<prior-version>.ts
   ```
5. Read the prior minor's doctor file:
   ```
   src/m365/spfx/commands/project/project-doctor/doctor-<prior-version>.ts
   ```
6. Read the compatibility matrix entry for the prior version in:
   ```
   src/m365/spfx/commands/SpfxCompatibilityMatrix.ts
   ```
7. Summarize findings to the user before proceeding: list what rule files will be created, what will change, and what test projects will be needed (see Phase 3 for naming conventions).

---

### Phase 2: Code Changes

All changes go on a dedicated branch (never commit to fork's `main`). Branch name: `add-spfx-<version>-support` (e.g. `add-spfx-1232-support`).

#### 1. Create `upgrade-<version>.ts`

Create `src/m365/spfx/commands/project/project-upgrade/upgrade-<version>.ts`.

- For **patch releases**: include only the `@microsoft/sp-*` package version rules, the ESLint/build devdep rules (`FN002022`, `FN002023`, `FN002030`, `FN002034`), and the `.yo-rc.json` version rule (`FN010001`). Do NOT copy one-time tooling changes (heft bumps, ESLint migration, flat config, SCSS) from the prior minor's file — those were version-specific.
- For **minor/major releases**: include all applicable rules, adding any new rule classes as needed.

Use object notation for all rules:
```typescript
new FN001001_DEP_microsoft_sp_core_library({ packageVersion: '1.23.2' })
```

Reference `upgrade-1.22.2.ts` (patch) or `upgrade-1.23.0.ts` (minor) as templates for the appropriate release type.

#### 2. Create `doctor-<version>.ts`

Create `src/m365/spfx/commands/project/project-doctor/doctor-<version>.ts`.

- For **patch releases**: copy `doctor-<prior-minor>.ts` verbatim — external dep ranges don't change.
- For **minor/major releases**: update ranges as specified in the release notes.

#### 3. Update `project-upgrade.ts`

Add `'<version>'` to the `supportedVersions` array in `src/m365/spfx/commands/project/project-upgrade.ts`, after the previous version entry.

#### 4. Update `project-doctor.ts`

Add `'<version>'` to the `supportedVersions` array in `src/m365/spfx/commands/project/project-doctor.ts`, after the previous version entry.

#### 5. Update `SpfxCompatibilityMatrix.ts`

Add a `'<version>'` entry in `src/m365/spfx/commands/SpfxCompatibilityMatrix.ts` after the prior version's block.

- For **patch releases**: copy the prior minor's entry exactly (same node, heft, yo ranges).
- For **minor/major releases**: update ranges per the release notes.

#### 6. Update `project-upgrade.mdx`

Update the "latest version" mention in `docs/docs/cmd/spfx/project/project-upgrade.mdx` (in the Remarks section) from the previous version to the new version.

#### 7. Add e2e tests to `project-upgrade.spec.ts`

Add a `//#region <prior-version>` block in `src/m365/spfx/commands/project/project-upgrade.spec.ts`, inserted before the existing `//#region superseded rules` section.

Add one test per test project type (see Phase 3 for the full list). Use placeholder count `0` initially — counts will be filled in after Phase 3:

```typescript
//#region <prior-version>
it('e2e: shows correct number of findings for upgrading webpart-react <prior-version> project to <version>', async () => {
  sinon.stub(command as any, 'getProjectRoot').callsFake(_ => path.join(process.cwd(), 'src/m365/spfx/commands/project/test-projects/spfx-<prior-version-short>-webpart-react'));
  await command.action(logger, { options: { toVersion: '<version>', preview: false, output: 'json' } } as any);
  const findings: FindingToReport[] = log[0];
  assert.strictEqual(findings.length, 0);
});
// ... one test per project type
//#endregion
```

Note: the region label and test project directory prefix use the **source** version (prior minor, shortened — e.g. `1230` for `1.23.0`), not the target.

---

### Phase 3: Test Project Scaffolding (⚠️ Human Operator Required — STOP)

> **Future toolchain note:** The SPFx CLI (`@microsoft/spfx-cli`) is currently in pre-release and is the planned replacement for the Yeoman generator (`@microsoft/generator-sharepoint`). See: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/toolchain/sharepoint-framework-cli
>
> Once the SPFx CLI reaches general availability, this phase will need to be updated:
> - Replace `yo @microsoft/generator-sharepoint` with `spfx create --template <type> --library-name spfx --component-name "HelloWorld" --spfx-version <version>`
> - Map existing project type names to the new template names (e.g. `webpart-react`, `extension-application-customizer`, `extension-fieldcustomizer-react`, `extension-formcustomizer-noframework`, `extension-formcustomizer-react`, `extension-listviewcommandset`, `ace-generic-card`)
> - Update the "Required tooling" section — `yo` and the generator package are replaced by `npm install -g @microsoft/spfx-cli`
> - The new CLI uses `--spfx-version` to target a version branch instead of requiring a specific generator version to be installed

**STOP here.** The following test projects must be manually scaffolded by the human operator using `yo @microsoft/generator-sharepoint` with the **prior minor's** SPFx packages installed. Claude cannot run interactive generators.

#### Required tooling

Verify before scaffolding:
- Node.js in the required range for the target version (check `SpfxCompatibilityMatrix.ts`)
- `yo` at a supported version
- `@microsoft/generator-sharepoint` at the **prior minor** version (e.g. `1.23.0`)
- `@rushstack/heft` at a supported version

#### Test projects to scaffold

Projects follow the naming pattern `spfx-<version-short>-<type>` (e.g. `spfx-1230-webpart-react`).

Standard project types (scaffold all that are relevant to the SPFx version):

| Directory suffix | Yeoman component type |
|---|---|
| `ace` | Adaptive Card Extension |
| `applicationcustomizer` | Application Customizer extension |
| `fieldcustomizer-react` | Field Customizer extension (React) |
| `formcustomizer-nolib` | Form Customizer (no framework) |
| `formcustomizer-react` | Form Customizer (React) |
| `listviewcommandset` | List View Command Set extension |
| `webpart-nolib` | Web Part (no framework) |
| `webpart-react` | Web Part (React) |
| `webpart-optionaldeps` | Web Part (React, with optional SPFx deps installed) |

For each project:
1. Scaffold with `yo @microsoft/generator-sharepoint` in a temp directory
2. Use solution name `spfx` and component name `HelloWorld` for consistency
3. Copy the scaffolded project (excluding `node_modules/`) into `src/m365/spfx/commands/project/test-projects/`

Also create one **target-version** project to verify doctor compatibility:

| Directory | Purpose |
|---|---|
| `spfx-<version-short>-webpart-react` | Scaffolded at the new version; used to verify `project-doctor` |

Signal Claude when all projects are committed and ready to proceed.

---

### Phase 4: Verification

After the human operator signals that test projects are ready:

#### 1. Build

```shell
npm run build
```

Fix any TypeScript compilation errors before proceeding.

#### 2. Run e2e tests and capture actual finding counts

```shell
npx mocha --no-config "dist/m365/spfx/commands/project/project-upgrade.spec.js" --reporter spec --timeout 30000 --grep "1\\.23\\.0.*1\\.23\\.2"
```

(Adjust the grep pattern to match the prior-version → target-version tests added in Phase 2.)

For each failing test, the actual count is in the error output (`expected 0 to equal N`). Update `assert.strictEqual(findings.length, <N>)` in the spec file, then re-run until all tests pass.

#### 3. Verify finding count consistency

Compare counts to the equivalent tests for the same project types in the prior patch release (if any). For a patch release, counts should differ by at most 1–2 findings (only the package version bumps added by this release). If counts differ dramatically, investigate before proceeding.

#### 4. Run doctor spec

```shell
npx mocha --no-config "dist/m365/spfx/commands/project/project-doctor.spec.js" --reporter spec --timeout 30000
```

#### 5. Run spfx doctor spec

```shell
npx mocha --no-config "dist/m365/spfx/commands/spfx-doctor.spec.js" --reporter spec --timeout 30000
```

All three spec runs must pass before opening the PR.

---

### Phase 5: PR

Follow the [contributing guidelines](https://pnp.github.io/cli-microsoft365/contribute/creating-the-pr).

1. Squash all branch commits into a single commit:
   - Commit message: `Adds support for SPFx v<version>. Closes #<issue>`
2. Push to the fork and open a PR targeting `pnp/cli-microsoft365` `main`.
3. PR title: `Adds support for SPFx v<version>`
4. PR description should include:
   - Summary of changes (new files, modified files)
   - List of new test projects and their scaffolding inputs (solution name, component name, framework, yo generator version)
   - Finding count consistency table comparing prior patch release counts to new counts per project type
   - Reference to the SPFx release notes URL
   - `Closes #<issue>`

**Note on Mocha output**: The default Mocha reporter (`scripts/ci-test-summary.cjs`) clears the screen using terminal escape sequences and produces empty output when stdout is piped. Always use `--no-config --reporter spec` when running tests to capture readable output.
