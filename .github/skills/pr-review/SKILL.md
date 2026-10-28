---
name: pr-review
description: >-
  This skill should be used when the user asks to "review a PR",
  "check this pull request", "PR review", "review pull request",
  "check PR against checklist", or needs to review a pull request
  against the project's PR checklist and coding standards.
---

# PR Review

Review pull requests against the project's PR checklist and coding standards. Produce a numbered list of findings for the user, then post inline review comments and a summary comment on the PR when approved.

## Prerequisites

- `gh` CLI authenticated with access to the repository.

## Workflow

### Phase 1: Gather PR Context

1. If no PR number or URL was provided, prompt the user for it. Do not proceed without one.
2. Extract the PR number. Accept formats: full URL, `#123`, or plain number.
3. Run the following to collect PR metadata and diff:

```bash
gh pr view <number> --json number,title,body,author,files,additions,deletions,baseRefName,headRefName,headRefOid,closingIssuesReferences
gh pr diff <number>
```

4. **Reviewing a specific commit**: `gh pr diff` always shows the latest PR state. When asked to review a specific commit (or an earlier state of the PR), resolve the PR's target branch and the base at the reviewed revision. Do not use the merge base with today's target branch: a historical PR commit that is already an ancestor of that branch would produce an empty diff.

```bash
repo=<base-owner>/<base-repo>   # repository containing the PR, not the contributor's fork
base_ref=$(gh api repos/$repo/pulls/<number> --jq '.base.ref')
first_pr_commit=$(gh api --paginate repos/$repo/pulls/<number>/commits --jq '.[].sha' | head -n 1)
base_sha=$(gh api repos/$repo/commits/$first_pr_commit --jq '.parents[0].sha')
git fetch <remote> $base_ref $base_sha <sha>   # <remote> points at the base repository
git diff $base_sha..<sha>
```

5. **Read the linked issue(s)** (`closingIssuesReferences`, or `Closes #<n>` in the body) with `gh issue view <n> --comments`. The issue is the specification: verify the PR implements what was agreed (option names, aliases, behavior, output).
6. Identify the type of change (new command, bug fix, refactor/migration, breaking change/removal, documentation, SPFx rules, docs site, infrastructure, etc.) from the PR title, body, linked issue, and changed files. This determines which checklist items and area-specific checklists apply.
7. **Read as widely as needed** to understand the change: callers, shared helpers the PR uses, related docs, release notes, and the linked issue. Wider reading informs findings but does not widen what can block (see [Scope policy](#scope-policy)).
8. **Build, test, and lint** when feasible, without modifying the user's working tree. Use a separate worktree:

```bash
git fetch <remote> pull/<number>/head:pr-<number>   # <remote> points at the base repository; this also works for fork PRs
review_sha=<sha>   # requested commit, or the headRefOid captured in step 3 for a latest-state review
git worktree add --detach ../pr-<number>-review $review_sha
cd ../pr-<number>-review && npm ci && npm run build
npx mocha --no-config --require source-map-support/register --timeout 10000 "dist/**/<changed-file>.spec.js"
npx eslint <changed files>
```

   `--no-config` prevents `.mocharc.json` from adding its repository-wide `dist/**/*.spec.js` pattern to the positional spec and turning a targeted run into the full suite.

   Remove the worktree when done (`git worktree remove ../pr-<number>-review`). If you did not build, test, or lint, say so explicitly in the findings and the review summary (for example, "Not verified by running tests or lint").

### Scope policy

Read widely, but only issues the PR **introduces or exposes** may be blocking findings.

In scope:

- Problems in new or changed lines.
- Unchanged code whose behavior changes because of the PR. For example: an unchanged `allowUnknownOptions()` that makes a removed option silently accepted; an unchanged helper (e.g. one that only matches `*.sharepoint.com`) that the PR newly relies on; a new `responseType` that interacts with error handling in `request.ts`; docs that still list options or aliases the PR dropped.
- New code that copies an existing bad pattern.

Out of scope:

- Pre-existing issues in code the PR doesn't change never block. Omit them, or list them under a non-blocking **Noticed, out of scope** heading, ideally suggesting a follow-up issue.
- For lines rewritten wholesale by a mechanical refactor or migration, flag only what the migration broke. Older problems on those lines follow the pre-existing rule.

**Heuristic**: "Would this still be a problem if the PR were reverted?" If yes, it is pre-existing and must not block.

### Phase 2: Review

Read `references/pr-checklist.md` for the full checklist. When the PR touches SPFx upgrade/doctor rules, the docs site, or telemetry/infrastructure, also read `references/area-checklists.md`.

Review the diff against **all four categories** below. Not every checklist item applies to every PR — evaluate applicability based on the change type identified in Phase 1.

#### A. Checklist Compliance

Walk through each applicable item in the checklist and verify it against the diff. Key items to check:

- **Single quotes** for strings in command files (not double quotes). For test files, this is less strictly enforced.
- **Command class naming**: `[Service][Command]Command` pattern.
- **Command name in `commands.ts`**: Must be placed in alphabetical order.
- **`force` option** on remove commands.
- **`async/await`** instead of `promise/then`.
- **`handleRejectedODataJsonPromise`** when `responseType: 'json'` is used.
- **`defaultProperties`** in list commands instead of conditional JSON output.
- **Custom Zod validation** for mutually exclusive options.
- **Verbose logging**: Every command should include at least one verbose log message that provides useful execution context.
- **User input escaped** in XML and URLs using `formatting.encodeQueryParameter()` or `formatting.escapeXml()`.
- **No `any` types**, no commented-out code.
- **Bug fixes** include a test for the fixed case.
- **Documentation** included where needed.

#### B. Code Quality

Review the implementation for:

- Logic errors, off-by-one mistakes, unhandled edge cases.
- Security issues (injection, unvalidated input, leaked or publicly exposed credentials and secrets).
- Performance concerns (unnecessary API calls, missing pagination, unbounded concurrency).
- TypeScript best practices (proper typing, no implicit any).

**Breaking changes and removals** (removed or renamed commands, options, aliases, or behavior):

- Release notes (`docs/docs/about/release-notes.mdx`) use accurate wording: *removed* vs *renamed* vs *deprecated*.
- Upgrade guidance is added for the target major version (`docs/docs/v<N>-upgrade-guidance.mdx`), explaining what changed and what users should do.
- No active usage of the removed option or command remains in source or current command documentation (e.g. `grep -rn -- '--<option>' src docs`). Preserve valid historical references in release notes, migration or upgrade guidance, and tests that verify the removed option is rejected.
- Removed options are actually rejected (check `allowUnknownOptions()` and schema strictness), not silently accepted.

**Refactors and migrations** (e.g. migrating a command to Zod) — verify behavior parity with the previous implementation:

- Short aliases are preserved (e.g. `-t`, `-m`, `-n`).
- Option sets keep the same semantics: exactly-one vs at-most-one; specifying none still fails when it did before.
- Option-set options use `customCode: 'optionSet'` in refinements so interactive prompting still works; `customCode: 'required'` is used only where an option is truly required.
- Enum and case handling is unchanged (e.g. use `zod.coercedEnum` for case-insensitive enums; check that transforms actually run).
- Unknown-options behavior is unchanged.
- Docs still match the options (names, aliases, required/optional, allowed values).

#### C. Test Quality

Check test files (`.spec.ts`) for:

- Coverage of happy path and error cases.
- Proper use of mocks and assertions.
- Tests that actually verify behavior (not just that code runs).
- Tests for command name, description not being `null`, and schema validation.
- Use `sinon.restore()` in `after()` to reset stubs/spies.
- Test names describe the validated behavior (e.g. `'retrieves file contents as string when asString is specified'`), not the implementation or a vague label.
- Assert exact URLs and request bodies (`opts.url === '<full url>'`, `assert.deepStrictEqual(postStub.lastCall.args[0].data, { ... })`), not `indexOf` or partial matches.
- Use `settingsNames` constants (`src/settingsNames.ts`) instead of hardcoded setting name strings.
- When the PR changes arguments passed to an API or helper (e.g. adds `permanent: true`), tests assert the changed argument is passed.

#### D. Documentation Quality

Check documentation files (`.mdx`) for:

- Every command needs a reference page at `docs/docs/cmd/<workload>/<command-name>.mdx`.
- The `.mdx` file name matches the command file name.
- Must include: title, description, usage, options table, examples, permissions, and response. A remarks section can be used for additional information, but is not mandatory.
- When a command has a response, include sample output for all four output formats: JSON, Markdown, text, and CSV.
- Each format's sample matches real output. Without `--query`, only `text` output is trimmed to `defaultProperties`; JSON includes all properties, CSV omits non-null object and array values, and Markdown omits object-valued properties, including arrays and `null`. A CSV or Markdown sample may therefore omit nested properties even when JSON includes them.
- Include at least 2 examples for the examples section, covering different option combinations.
- Import and use `<Global />` for standard CLI options.
- Examples should use `m365` prefix and long option names (not short aliases).
- Document the minimum required permissions that allow success with any option combination.
- Docs must build without warnings.
- When importing components in the docs, use absolute imports from `@site/src/components/` instead of relative paths.
- When importing global options in the docs, use a relative import from `../_global.mdx`.

### Phase 3: Present Findings

Compile all in-scope findings into a **numbered list** and present it to the user in chat. Each finding must include:

1. **File and line reference**, linked as a blob URL at the reviewed commit (PR `headRefOid` or the specific SHA) with line anchors: `[<path>:<line>](https://github.com/<owner>/<repo>/blob/<sha>/<path>#L<start>-L<end>)`. Use `#L<n>` for a single line.
2. **Category**: Checklist | Code Quality | Test Quality | Documentation.
3. **Severity**: Error (must fix) | Warning (should fix) | Suggestion (nice to have).
4. **Description**: What the issue is and why it matters.
5. **Suggestion**: What to change (be specific).

**Severity calibration**:

- **Error**: bugs, security issues (e.g. leaked or publicly exposed secrets, injection), and regressions in behavior users rely on (e.g. dropped aliases, broken interactive prompting, removed options silently accepted) are always Errors.
- **Warning**: checklist violations and issues that should be fixed but don't break users.
- **Suggestion**: style, readability, and optional improvements.

If there are **zero findings**, state that the PR looks good.

Sort findings: Errors first, then Warnings, then Suggestions. List out-of-scope observations separately under **Noticed, out of scope** (non-blocking). State whether build, tests, and lint were run and their results.

### Phase 4: Post Comments

Always ask the user for confirmation before posting any comments on the PR. Present the planned comments and wait for explicit approval. Then:

1. **Submit a review** with inline comments and the appropriate verdict:
   - **Any Error** → `REQUEST_CHANGES`
   - **Only Warnings and/or Suggestions** → `COMMENT`
   - **Zero findings** → `APPROVE`

   Out-of-scope observations never affect the verdict.

Post inline comments and the summary as a single review submission. Always target the actual line where the issue exists.

**Line number accuracy**: When posting inline comments, verify line numbers against the actual diff hunk content. The line number refers to the **new file** line number (RIGHT side). Count lines in the diff hunk carefully — off-by-one errors cause comments to land on the wrong line and suggestion blocks to replace the wrong code. Inline comments can only target lines in the diff; for issues in unchanged code exposed by the PR, include them in the summary.

```bash
gh api repos/{owner}/{repo}/pulls/{number}/reviews \
  --method POST \
  --input - <<'EOF'
{
  "commit_id": "<reviewed sha>",
  "event": "APPROVE, COMMENT, or REQUEST_CHANGES",
  "body": "<summary>",
  "comments": [
    {
      "path": "<file>",
      "line": <line>,
      "side": "RIGHT",
      "body": "<comment>"
    }
  ]
}
EOF
```

The `comments` array can contain multiple entries. Omit it entirely when approving with zero findings.

#### Comment Tone Guidelines

These comments appear on a public repository and represent the user responding to a contributor's work. Every comment must be:

- **Constructive**: Frame issues as suggestions, not demands. Use "Consider..." or "It might be worth..." instead of "You must..." or "This is wrong."
- **Specific**: Reference exact lines, show the expected code, explain *why* a change matters.
- **Appreciative**: Acknowledge good work. If the PR is mostly solid, say so. Contributors volunteer their time.
- **Brief**: One or two sentences per inline comment. Save detail for the summary.
- **Professional**: No sarcasm, no passive-aggressive phrasing, no exclamation marks on criticism.

For inline comments, use this format, where Category is one of Checklist, Code Quality, Test Quality, or Documentation:
```
[Category | Severity] Description.

Suggestion: `<code or guidance>`
```

When confident a specific code change is correct, include a GitHub suggestion block to make it easy for the contributor to apply:

````
```suggestion
<corrected code>
```
````

Only use suggestion blocks when the fix is unambiguous and self-contained. Do not guess.

For the summary comment, use this format:
```markdown
## PR Review Summary

**PR**: #<number> - <title>
**Author**: @<author>

### Overview
<1-2 sentence assessment>

### Findings
<numbered list of findings with severity>

### Noticed, out of scope
<optional: non-blocking pre-existing issues, ideally with a suggested follow-up issue; omit the section if empty>

### Checklist Coverage
<brief note on which checklist categories were verified, and whether build/tests/lint were run>

---
*Reviewed against the [PR checklist](docs/docs/contribute/pr-checklist.mdx).*
```

If the PR has zero issues, the summary should simply acknowledge the quality of the work and approve.

## Additional Resources

### Reference Files

- **`references/pr-checklist.md`** — Full PR checklist with all items organized by category (General, Coding Standards, Documentation).
- **`references/area-checklists.md`** — Short checklists for SPFx upgrade/doctor rules, docs site changes, and telemetry/infrastructure changes.

### Reference Commands

When providing feedback, point contributors to well-implemented commands as examples. Use these as reference implementations:

- **`src/m365/spo/commands/list/list-list.ts`** — Example of a list command with `defaultProperties` and proper text output.
- **`src/m365/spo/commands/web/web-get.ts`** — Example of a get command with verbose logging and proper error handling.
- **`src/m365/entra/commands/user/user-get.ts`** — Example of a command using Zod schema validation.
