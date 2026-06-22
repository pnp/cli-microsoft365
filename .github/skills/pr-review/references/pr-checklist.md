# PR Checklist

## General guidelines

- Ensure the build passes.
- Achieve 100% code coverage.
- The submission should match the specification.
- Maintain a single commit, or squash multiple commits.
- Use single quotes `' '` for strings.
- If submitting a sample, ensure it is properly formatted and indented.

## Coding standards

- Command options should follow the naming convention (kebab-case for CLI flags).
- The command should have a correct name.
- The command name added to `commands.ts` should be placed so that commands are sorted alphabetically.
- The command class is named following the pattern `[Service][Command]Command`. For example, `SpoWebRemoveCommand`.
- Verify the command works as expected.
- List commands must have readable output in `text` mode, with each item fitting in one row of 130 characters preferably.
- Avoid commented-out code and usage of `any` types, preferring specific types.
- Remove commands should include a `force` option.
- For bug fixes, include a test for the fixed use case.
- Avoid unnecessary retrieval of form digest.
- Handle failed promises properly when `responseType: 'json'` is used by using `handleRejectedODataJsonPromise`.
- Escape user input in XML and URLs.
- Verbose and debug outputs are logged to stdErr (`logger.logToStderr` instead of `logger.log`).
- Do not do conditional output in JSON output mode; use `defaultProperties` for defining default properties.
- For commands with multiple options where the user is required to choose one, define these options using a custom Zod validation.
- Use `async/await` instead of `promise/then`.
- When working with `spo` commands, use `GetFileByServerRelativePath` and `GetFolderByServerRelativePath` API endpoint instead of `GetFileByServerRelativeUrl` and `GetFolderByServerRelativeUrl`.
- `npm test` must pass without errors.

## Documentation

- Include an `mdx` help page.
- Reference the `mdx` help page in the sidebar navigation.
- Start all code samples with `m365`.
- Ensure samples use long names of options rather than short ones.
- Include the marker to incorporate global options rather than listing them explicitly.
- Check for no warnings when building docs (lines that begin with `WARNING - `).
- If there is an option modifying the output, include responses for both default and modified output.
