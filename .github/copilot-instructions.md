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

## Command Implementation

- Command class should extend the appropriate workload-specific base class (`SpoCommand`, `GraphCommand`, `AzmgmtCommand`, etc.) when applicable; otherwise, it may extend `Command` directly.
- Class name must follow the `[Service][Entity][Action]Command` pattern (e.g., `SpoGroupGetCommand`).
- Must implement `name`, `description`, and `commandAction()` at minimum.
- Options must be defined using **Zod schemas** with `z.strictObject({ ...globalOptionsZod.shape, ... })` or `globalOptionsZod.strict()`.
- Option names must be camelCase; apply aliases with the schema `.alias('x')` method (for example, `.alias('u')` for `--webUrl` where appropriate).
- Commands that delete/remove resources must include a `--force` option.
- Use `async/await`, never `.then()` chains.
- Use `logger.logToStderr()` for verbose/debug messages, never `logger.log()`.
- Add verbose log messages where they provide useful execution context; every command requires at least one verbose logging statement.
- Never format output conditionally based on `--output json`; use `defaultProperties` for text field filtering.
- All `commands.ts` files are sorted alphabetically by command name; new commands must be added in the correct order.

## API Usage

- Use `request.ts` wrapper for all HTTP calls, never call APIs directly.
- Use `handleRejectedODataJsonPromise()` for JSON response error handling.
- Escape user input in XML payloads and URL parameters using `formatting.encodeQueryParameter()` or `formatting.escapeXml()`.
- Avoid unnecessary form digest retrieval.

## Testing

- Test file must be alongside the command file with `.spec.ts` extension.
- Use Mocha (`describe`/`it`/`before`/`beforeEach`/`afterEach`/`after`) + Sinon (stubs/spies).
- Standard test setup must stub: `auth.restoreAuth`, `telemetry.trackEvent`, `pid.getProcessName`, `session.getId`.
- Must set `auth.connection.active = true` in `before()`.
- Must initialize `commandInfo` and `commandOptionsSchema` in `before()` for Zod-based commands.
- Must include tests for: command name matches constant, description is not null, schema validation (valid and invalid options), success paths, error handling paths.
- Options in `command.action()` calls must be parsed through `commandOptionsSchema.parse()`, not passed as raw objects.
- Never use `as any` to bypass type checking unless absolutely necessary for error path testing.
- Use `sinon.restore()` in `after()` to reset stubs/spies.

## Documentation

- Every command needs a reference page at `docs/docs/cmd/<workload>/<command-name>.mdx`.
- The `.mdx` file name matches the command file name.
- Must include: title, description, usage, options table, examples, permissions, and response.
- The remarks section can be used for additional information but is not mandatory.
- When a command has a response, include sample output for all four output formats: JSON, Markdown, text, and CSV.
- Include at least 2 examples for the examples section, covering different option combinations.
- Import and use `<Global />` for standard CLI options.
- Examples should use `m365` prefix and long option names (not short aliases).
- Document minimum required permissions that allow success with any option combination.
- Docs must build without warnings.
- When importing components in the docs, use absolute imports from `@site/src/components/` instead of relative paths.
- When importing global options in the docs, use a relative import from `../_global.mdx`.

## Code Quality

- No `any` types — use proper TypeScript interfaces/types.
- No commented-out code.
- Single quotes for all strings, where possible.
- No unused imports or variables.
- Do not quote property names in JSON/object literals unless the property name requires it (e.g., contains special characters or hyphens).
- Follow existing patterns in neighboring command files for consistency.
- Custom ESLint rules are enforced: command class naming, command name dictionary, no deprecated API usage.
- New command name words may require adding to the ESLint dictionary in `eslint.config.mjs`.

## Developer Workflows

- **Build**: `npm run build` (TypeScript compile + command metadata generation)
- **Test**: `npm test` (runs version check, lint, and Mocha tests)
- **Lint**: `npm run lint` (uses custom ESLint rules from `eslint-rules/`)
- **Watch**: `npm run watch` (TypeScript in watch mode)
- **Symlink for local CLI**: `npm link` (after build)
- **Node version**: Must be 24 (see `scripts/check-version.js`)
- **Docs**: Docusaurus site in `docs/` (config: `docs/docusaurus.config.ts`)

## Project-Specific Notes

- **Command registration**: New commands must be discoverable for `write-all-commands.js` to pick up.
- **SPFx support**: Special logic for SPFx project upgrades and compatibility checks in `src/m365/spfx/`.
- **Output**: Prefer returning objects/arrays; formatting handled by CLI core.

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
