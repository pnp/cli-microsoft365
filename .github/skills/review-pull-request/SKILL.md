---
name: review-pull-request
description: "Review a pull request for CLI for Microsoft 365. Use when reviewing a PR, providing PR feedback, checking code quality, or validating contributions against project conventions."
---

# Pull Request Review Guidelines for CLI for Microsoft 365

When reviewing a PR, verify each of the following areas. For detailed coding conventions (command implementation, API usage, testing, documentation, and code quality), refer to `.github/copilot-instructions.md`.

## PR Structure

- PR title should be descriptive (e.g., "Adds 'spo site get' command")
- PR should reference the issue it addresses (e.g., "Closes #1234")
- PR should target the `main` branch only
- New command PRs should include changes to: command file, spec file, `commands.ts`, docs `.mdx` file, and `docs/src/config/sidebars.ts`

## Review Checklist

- For bug fixes: include a test that reproduces the original bug
