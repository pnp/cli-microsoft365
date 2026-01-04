# Copilot Instructions for CLI for Microsoft 365

## Project Overview
- This is a cross-platform CLI tool to manage Microsoft 365 tenants and SharePoint Framework (SPFx) projects.
- Written in TypeScript, targeting Node.js LTS (see `package.json`, `tsconfig.json`).
- Main entry: `src/index.ts` → `src/cli/cli.ts` (command parsing/execution).
- Commands are modular, organized under `src/m365/` by workload (e.g., `spo`, `entra`, `teams`).
- Each command is a class extending `Command` (`src/Command.ts`).
- Command metadata and options: see `src/cli/CommandInfo.ts`, `src/cli/CommandOptionInfo.ts`.
- Command discovery/build: see `scripts/write-all-commands.js` (generates `allCommands.json`).

## Key Patterns & Conventions
- **Global options**: All commands accept `--output`, `--query`, `--debug`, `--verbose` (see `src/GlobalOptions.ts`).
- **Authentication**: Managed via `src/Auth.ts`, supports multiple auth types (Device Code, Certificate, Secret, Managed Identity, etc.).
- **HTTP requests**: Use the wrapper in `src/request.ts` (Axios-based, with logging/debug hooks).
- **Telemetry**: Centralized in `src/telemetry.ts`.
- **Validation**: Use Zod schemas (`zod` import, see `src/Command.ts`).
- **Prompting**: Use utilities in `src/utils/prompt.ts` for interactive input.
- **Config**: User settings/configuration via `configstore` (see `src/config.ts`).
- **Completion**: Shell completion logic/scripts in `scripts/Register-CLIM365Completion.ps1` and `scripts/Test-CLIM365Completion.ps1`.

## Developer Workflows
- **Build**: `npm run build` (TypeScript compile + command metadata generation)
- **Test**: `npm test` (runs version check, lint, and Mocha tests)
- **Lint**: `npm run lint` (uses custom ESLint rules from `eslint-rules/`)
- **Watch**: `npm run watch` (TypeScript in watch mode)
- **Symlink for local CLI**: `npm link` (after build)
- **Node version**: Must be 24 (see `scripts/check-version.js`)
- **Docs**: Docusaurus site in `docs/` (config: `docs/docusaurus.config.ts`)

## Project-Specific Notes
- **Command structure**: Each command is a class, not a function. Use inheritance from `Command`.
- **Command registration**: New commands must be discoverable for `write-all-commands.js` to pick up.
- **SPFx support**: Special logic for SPFx project upgrades and compatibility checks in `src/m365/spfx/`.
- **Output**: Prefer returning objects/arrays; formatting handled by CLI core.
- **No direct file/console output in commands**: Use provided logger and output mechanisms.
- **Testing**: Mocha-based, see test files alongside source (e.g., `*.spec.ts`). 100% code coverage.
- **Custom ESLint rules**: See `eslint-rules/` and `eslint-plugin-cli-microsoft365` in `package.json`.

## Integration & External Dependencies
- **Microsoft Graph, SharePoint REST, etc.**: Use `request.ts` for all HTTP calls.
- **Authentication**: Uses `@azure/msal-node` and related packages.
- **Telemetry**: Application Insights via `applicationinsights` package.
- **Docs**: Docusaurus, see `docs/` and `docs/docusaurus.config.ts`.

## Examples
- Add a new command: Get the latest information about building commands from [the docs](../docs/docs/contribute/new-command/build-command-logic.mdx). Create a class in the appropriate `src/m365/<workload>/commands/` folder, extend `Command`, implement `commandAction`, register options, and ensure it is discoverable. Create a test file next to the command file. Create a reference page in the documentation in `docs/docs/cmd/<workload>`. The reference page file name is the same as the command file name, but with the `.mdx` extension.
- Add a global option: update `src/GlobalOptions.ts` and propagate to command parsing in `src/cli/cli.ts`.

## References
- [Main README](../README.md)
- [Contributing Guide](../docs/docs/contribute/contributing-guide.mdx)
- [Docs site](https://pnp.github.io/cli-microsoft365/)
