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
gh pr view <number> --json number,title,body,author,files,additions,deletions,baseRefName,headRefName
gh pr diff <number>
```

4. Identify the type of change (new command, bug fix, documentation, refactor, etc.) from the PR title, body, and changed files. This determines which checklist items apply.

### Phase 2: Review

Read `references/pr-checklist.md` for the full checklist.

Review the diff against **all three categories** below. Not every checklist item applies to every PR — evaluate applicability based on the change type identified in Phase 1.

#### A. Checklist Compliance

Walk through each applicable item in the checklist and verify it against the diff. Key items to check:

- **Single quotes** for strings (not double quotes).
- **Command class naming**: `[Service][Command]Command` pattern.
- **Command name in `commands.ts`**: Must be placed in alphabetical order.
- **`force` option** on remove commands.
- **`async/await`** instead of `promise/then`.
- **`handleRejectedODataJsonPromise`** when `responseType: 'json'` is used.
- **`logger.logToStderr`** for verbose/debug output (not `logger.log`).
- **`defaultProperties`** instead of conditional JSON output.
- **Custom Zod validation** for mutually exclusive options.
- **`GetFileByServerRelativePath`/`GetFolderByServerRelativePath`** in `spo` commands (not `...Url` variants).
- **No `any` types**, no commented-out code.
- **User input escaped** in XML and URLs.
- **Bug fixes** include a test for the fixed case.
- **Documentation** included where needed.

#### B. Code Quality

Review the implementation for:

- Logic errors, off-by-one mistakes, unhandled edge cases.
- Security issues (injection, unvalidated input, leaked credentials).
- Performance concerns (unnecessary API calls, missing pagination).
- TypeScript best practices (proper typing, no implicit any).

#### C. Test Quality

Check test files (`.spec.ts`) for:

- Coverage of happy path and error cases.
- Proper use of mocks and assertions.
- Tests that actually verify behavior (not just that code runs).
- Tests for command name, description not being `null`, and schema validation.
- Use `sinon.restore()` in `after()` to reset stubs/spies.

#### D. Documentation Quality

Check documentation files (`.mdx`) for:

- Every command needs a reference page at `docs/docs/cmd/<workload>/<command-name>.mdx`.
- The `.mdx` file name matches the command file name.
- Must include: title, description, usage, options table, examples, permissions, and response. A remarks section can be used for additional information, but is not mandatory.
- When a command has a response, include sample output for all four output formats: JSON, Markdown, text, and CSV.
- Include at least 2 examples for the examples section, covering different option combinations.
- Import and use `<Global />` for standard CLI options.
- Examples should use `m365` prefix and long option names (not short aliases).
- Document the minimum required permissions that allow success with any option combination.
- Docs must build without warnings.
- When importing components in the docs, use absolute imports from `@site/src/components/` instead of relative paths.
- When importing global options in the docs, use a relative import from `../_global.mdx`.

### Phase 3: Present Findings

Compile all findings into a **numbered list** and present it to the user in chat. Each finding must include:

1. **File and line reference** (linked).
2. **Category**: Checklist | Code Quality | Test Quality.
3. **Severity**: Error (must fix) | Warning (should fix) | Suggestion (nice to have).
4. **Description**: What the issue is and why it matters.
5. **Suggestion**: What to change (be specific).

If there are **zero findings**, state that the PR looks good.

Sort findings: Errors first, then Warnings, then Suggestions.

### Phase 4: Post Comments

Always ask the user for confirmation before posting any comments on the PR. Present the planned comments and wait for explicit approval. Then:

1. **Submit a review** with inline comments and the appropriate verdict:
   - **Zero findings** → `APPROVE`
   - **Any findings** (error, warning, or suggestion) → `REQUEST_CHANGES`

Post inline comments and the summary as a single review submission. Always target the actual line where the issue exists.

**API limitation**: The GitHub API can only target lines within diff hunk ranges (changed lines + surrounding context lines). If the target line falls in a gap between hunks, do **not** post it on a nearby line. Instead, include that finding in the summary (the review body) with the file path and line number. Only use `suggestion` blocks for findings posted on the exact target line.

**Line number accuracy**: When posting inline comments, verify line numbers against the actual diff hunk content. The line number refers to the **new file** line number (RIGHT side). Count lines in the diff hunk carefully — off-by-one errors cause comments to land on the wrong line and suggestion blocks to replace the wrong code.

```bash
gh api repos/{owner}/{repo}/pulls/{number}/reviews \
  --method POST \
  --input - <<'EOF'
{
  "event": "APPROVE or REQUEST_CHANGES",
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

For inline comments, use this format:
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

### Checklist Coverage
<brief note on which checklist categories were verified>

---
*Reviewed against the [PR checklist](docs/docs/contribute/pr-checklist.mdx).*
```

If the PR has zero issues, the summary should simply acknowledge the quality of the work and approve.

## Additional Resources

### Reference Files

- **`references/pr-checklist.md`** — Full PR checklist with all items organized by category (General, Coding Standards, Documentation).

### Reference Commands

When providing feedback, point contributors to well-implemented commands as examples. Use these as reference implementations:

- **`src/m365/spo/commands/list/list-list.ts`** — Example of a list command with `defaultProperties` and proper text output.
- **`src/m365/spo/commands/web/web-get.ts`** — Example of a get command with verbose logging and proper error handling.
- **`src/m365/entra/commands/user/user-get.ts`** — Example of a command using Zod schema validation.
