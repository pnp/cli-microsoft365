---
name: update-release-notes
description: 'Update release notes based on the latest changes from the current branch. Use when: "update release notes", "add to release notes", "write release notes", "document changes in release notes", "sync release notes with branch".'
disable-model-invocation: true
---

# Update Release Notes

Update `docs/docs/about/release-notes.mdx` based on the changes in the current branch compared to main.

## Procedure

### 1. Gather Branch Changes

Run:

Get commit hash of latest release notes update

```shell
git log --grep="Updates release notes" -n 1 --format="%H"
```

Use the commit hash

```shell
git log <commit_hash>..HEAD --oneline --no-decorate
```

If the branch has no commits ahead of main, inform the user and stop.

For each commit, extract:
- The commit message (which follows the PR title convention)
- The linked issue/PR number (from `Closes #NNN` or `#NNN` in the message)

### 2. Read Current Release Notes

Read `docs/docs/about/release-notes.mdx` to find the current unreleased version section at the top of the file. The unreleased version heading looks like `## vX.Y.Z` (no link), while released versions have links: `## [vX.Y.Z](url)`.

If there is no unreleased version section (the first heading is a released version with a link), inform the user and stop. Do not create new version headings.

### 3. Classify Each Change

Classify each commit into one of these categories based on the commit message:

| Commit message pattern | Category |
|---|---|
| Starts with `Adds '...' command` | **New command** |
| Starts with `Fixes` | Change (fix) |
| Starts with `Extends`, `Enhances` | Change (enhancement) |
| Starts with `Migrates` | Change (migration) |
| Everything else (updates, removes, adds support, etc.) | Change (other) |

### 4. Analyze New Commands

For each new command, determine:

1. **Command name**: extract from the commit message (e.g., `spo site list`)
2. **Workload**: derive from the command name prefix:
   - `spo` → SharePoint
   - `spe` → SharePoint Embedded
   - `spp` → SharePoint Premium
   - `entra` → Entra ID
   - `teams` → Teams
   - `outlook` → Outlook
   - `viva` → Viva
   - `external` → External
   - `flow` → Flow
   - `Pa` → Power Apps
   - `pp` → Power Platform
   - `planner` → Planner
   - `purview` → Purview
   - `booking` → Booking
   - `onedrive` → OneDrive
3. **Doc path**: find the matching `.mdx` file under `docs/docs/cmd/` using the command name segments
4. **Description**: read the doc file's first paragraph or use the commit message description

### 5. Format Entries

#### New command entry

```markdown
- [command name](../cmd/<workload>/<subgroup>/command.mdx) - description [#issue](https://github.com/pnp/cli-microsoft365/issues/issue)
```

- Command name is the full command (e.g., `spo site list`)
- Description starts lowercase
- Issue link uses `issues/` for issues, `pull/` for PRs

#### Change entry

```markdown
- verb-past-tense description [#issue](https://github.com/pnp/cli-microsoft365/issues/issue)
```

Use past tense for the verb:
- `Adds` → `added`
- `Fixes` → `fixed`
- `Extends`/`Enhances` → `enhanced`
- `Migrates` → `migrated`
- `Updates` → `updated`
- `Removes` → `removed`

When the change relates to a specific command, link to its doc page inline:
```markdown
- fixed [command name](../cmd/path.mdx) command [#issue](url)
```

### 6. Insert Into Release Notes

Insert entries into the **current unreleased version section** (the first `## vX.Y.Z` heading without a link).

#### New commands

- If a `### New commands` section exists, add entries there
- If it doesn't exist and there are new commands, create it right after the version heading
- Group new commands under bold workload headers (`**SharePoint:**`, `**Outlook:**`, etc.)
- If a workload group already exists, append to it
- If a workload group doesn't exist, add it in alphabetical order among existing groups
- Separate workload groups with an empty line

#### Changes

- If a `### Changes` section exists, add entries there
- If it doesn't exist and there are changes, create it after `### New commands` (or after the version heading if there are no new commands)
- Add new entries at the end of the existing changes list

### 7. Verify

After editing, read back the modified section to verify:
- Entries are in the correct sections
- Links use relative paths (`../cmd/...`)
- Issue/PR numbers are linked correctly
- Description format matches existing entries (lowercase start, past tense)
- No duplicate entries were added
