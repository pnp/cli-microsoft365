---
name: refresh-clim365-skill
description: >-
  Regenerate the clim365 skill's command index after adding, removing, or
  renaming CLI for Microsoft 365 commands. Use when asked to "refresh the
  clim365 skill", "update the command index", "regenerate commands.txt",
  or after shipping a new command.
---

# Refresh the clim365 Skill Command Index

Regenerate `skills/clim365/references/commands.txt` so the clim365 skill can discover all current commands.

## Prerequisites

- `allCommandsFull.json` must be up to date. If you just added or changed commands, run `npm run build` first — the build generates this file.

## Workflow

1. **STOP — Confirm `allCommandsFull.json` is current.** If the user just added a command, ask whether they've run `npm run build`. If not, run it now.
2. Run the generation script from the repo root:
   ```sh
   node .github/skills/refresh-clim365-skill/references/write-skill-commands.js
   ```
3. Verify the output:
   ```sh
   wc -l skills/clim365/references/commands.txt
   ```
   The line count should match the number of commands in `allCommandsFull.json`.
4. Done. The updated `commands.txt` will be picked up by the clim365 skill on next use.
