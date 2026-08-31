---
name: new-command
description: >-
  This skill should be used when the user asks to "build a new command",
  "create a command", "implement a command", "add a new CLI command",
  or needs to build a new command for CLI for Microsoft 365 from a
  GitHub issue spec. It covers the full workflow: command logic,
  unit tests, documentation, sidebar registration, and PR checklist
  verification.
---

# Building a New Command for CLI for Microsoft 365

Build a complete, production-ready CLI command from a GitHub issue spec. The workflow produces four artifacts: command implementation, unit tests, documentation page, and sidebar registration — then verifies everything against the PR checklist.

## Prerequisites

A GitHub issue containing the command spec (name, description, options, examples, API details). If no issue is provided, **STOP — ask the user for the issue URL or spec before proceeding.**

## Workflow

Execute each phase in order. Do not skip phases.

### Phase 1: Parse the Spec

1. Read the GitHub issue thoroughly.
2. Extract: command name, description, service/workload, options (required/optional, types, aliases, allowed values, option sets), API endpoints used, example usage, and expected response shape.
3. **STOP — Verify API details are complete.** The spec must include the full API endpoint(s), HTTP method(s), request payloads, and response shapes. If any of these are missing, **ask the user** to provide them or point to API documentation. **NEVER fabricate or infer API request/response shapes** — even if similar commands exist in the codebase.
4. Identify the base class. Look at existing commands in `src/m365/<service>/commands/` to determine which base class to extend (`SpoCommand`, `GraphCommand`, `GraphApplicationCommand`, `AzmgmtCommand`, etc.).
5. Check that every word in the command name exists in the dictionary in `eslint.config.mjs`. If a word is missing, add it to the `dictionary` array (keep alphabetical order).

### Phase 2: Implement the Command

Create `src/m365/<service>/commands/<noun>/<noun>-<verb>.ts`.

#### Structure

```typescript
import { globalOptionsZod } from '../../../../Command.js';
import { z } from 'zod';
import { Logger } from '../../../../cli/Logger.js';
import commands from '../../commands.js';
import <BaseCommand> from '../../../base/<BaseCommand>.js';
import request, { CliRequestOptions } from '../../../../request.js';
// additional imports as needed

// Enums for options with predefined values
// enum Foo { Bar = 'bar', Baz = 'baz' }

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  // command-specific options
});
declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}

class <Service><Noun><Verb>Command extends <BaseCommand> {
  public get name(): string {
    return commands.<NOUN>_<VERB>;
  }

  public get description(): string {
    return '<description from spec>';
  }

  public get schema(): z.ZodType {
    return options;
  }

  // getRefinedSchema — only if option sets or cross-field validation needed

  public async commandAction(logger: Logger, args: CommandArgs): Promise<void> {
    try {
      if (this.verbose) {
        await logger.logToStderr(`<Verbose message>...`);
      }

      const requestOptions: CliRequestOptions = {
        url: `<endpoint>`,
        headers: { accept: 'application/json;odata.metadata=none' },
        responseType: 'json'
      };

      const result = await request.get<any>(requestOptions);
      await logger.log(result);
    }
    catch (err: any) {
      this.handleRejectedODataJsonPromise(err);
    }
  }
}

export default new <Service><Noun><Verb>Command();
```

#### Key rules

- Class name: `<Service><Noun><Verb>Command` in PascalCase.
- Options: use `z.strictObject` spreading `globalOptionsZod.shape`.
- Aliases: `.alias('x')` on the Zod property.
- Enums: `zod.coercedEnum(MyEnum)` for case-insensitive matching. Import `{ zod }` from `../../../../utils/zod.js`.
- Validation: Zod refinements on properties (`.refine()`), not custom validate methods.
- URL validation for SharePoint: `.refine(url => validation.isValidSharePointUrl(url) === true, { error: '...' })`.
- Option sets: implement `getRefinedSchema(schema)` returning `schema.refine(...)`.
- Async/await only — no `.then()`.
- Verbose/debug logging → `logger.logToStderr`.
- Error handling → `this.handleRejectedODataJsonPromise(err)`.
- SPO file/folder endpoints: use `GetFileByServerRelativePath` / `GetFolderByServerRelativePath`.
- Remove commands: include a `force` option and confirmation prompt using `cli.handleMultipleResultsFound` or `cli.promptForConfirmation`.
- No `any` types (except the catch clause). Use specific interfaces/types.
- No commented-out code.

#### Register the command name

Add the command constant to `src/m365/<service>/commands.ts`, keeping groups alphabetically sorted:

```typescript
export default {
  // ...existing commands...
  <NOUN>_<VERB>: `${prefix} <noun> <verb>`,
  // ...
};
```

### Phase 3: Write Unit Tests

Create `src/m365/<service>/commands/<noun>/<noun>-<verb>.spec.ts`.

#### Skeleton

```typescript
import assert from 'assert';
import sinon from 'sinon';
import auth from '../../../../Auth.js';
import { CommandError } from '../../../../Command.js';
import { cli } from '../../../../cli/cli.js';
import { CommandInfo } from '../../../../cli/CommandInfo.js';
import { Logger } from '../../../../cli/Logger.js';
import { telemetry } from '../../../../telemetry.js';
import { pid } from '../../../../utils/pid.js';
import { session } from '../../../../utils/session.js';
import { sinonUtil } from '../../../../utils/sinonUtil.js';
import request from '../../../../request.js';
import commands from '../../commands.js';
import command, { options as commandOptionsSchema } from './<noun>-<verb>.js';

describe(commands.<NOUN>_<VERB>, () => {
  let log: any[];
  let logger: Logger;
  let loggerLogSpy: sinon.SinonSpy;
  let commandInfo: CommandInfo;

  before(() => {
    sinon.stub(auth, 'restoreAuth').resolves();
    sinon.stub(telemetry, 'trackEvent').resolves();
    sinon.stub(pid, 'getProcessName').returns('');
    sinon.stub(session, 'getId').returns('');
    auth.connection.active = true;
    commandInfo = cli.getCommandInfo(command);
  });

  beforeEach(() => {
    log = [];
    logger = {
      log: async (msg: string) => { log.push(msg); },
      logRaw: async (msg: string) => { log.push(msg); },
      logToStderr: async (msg: string) => { log.push(msg); }
    };
    loggerLogSpy = sinon.spy(logger, 'log');
  });

  afterEach(() => {
    sinonUtil.restore([
      request.get,
      request.post,
      request.put,
      request.patch,
      request.delete
      // restore only the HTTP methods actually stubbed
    ]);
  });

  after(() => {
    sinon.restore();
    auth.connection.active = false;
  });

  it('has the correct name', () => {
    assert.strictEqual(command.name, commands.<NOUN>_<VERB>);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  // Validation tests — one pass and one fail per validation rule
  // Option set tests — valid combos and invalid combos
  // commandAction tests — one per branch/code path
  // API error test
});
```

#### Required test categories

1. **Name and description** — always.
2. **Validation** — each Zod refinement tested for pass and fail using `commandOptionsSchema.safeParse(...)`.
3. **Option sets** — valid single option, invalid multiple options, missing required option.
4. **Command action** — one test per logical branch. Stub `request.get`/`post`/etc. with `callsFake` matching URL patterns.
5. **Error handling** — stub request to reject, assert `CommandError`.
6. **Coverage** — every `if`, `switch`, `catch` branch hit. Target 100% code and branch coverage.

#### Run tests

```bash
npm test
```

Check coverage in `coverage/lcov-report/index.html`. If coverage is below 100% on the new command file, add tests for missed branches.

### Phase 4: Write Documentation

Create `docs/docs/cmd/<service>/<noun>/<noun>-<verb>.mdx`.

#### Template

````mdx
import Global from '../../_global.mdx';
import Tabs from '@theme/Tabs';
import TabItem from '@theme/TabItem';

# <service> <noun> <verb>

<Description from spec>

## Usage

```sh
m365 <service> <noun> <verb> [options]
```

## Options

```md definition-list
`-<alias>, --<option> <<option>>`
: <Description>. <Constraints>.

`--<optionalOption> [<optionalOption>]`
: <Description>.
```

<Global />

## Permissions

<!-- Generate with: node ./scripts/generate-docs-permissions.mjs -->

<Tabs>
  <TabItem value="Delegated">

  | Resource   | Permissions |
  |------------|-------------|
  | ...        | ...         |

  </TabItem>
  <TabItem value="Application">

  | Resource   | Permissions |
  |------------|-------------|
  | ...        | ...         |

  </TabItem>
</Tabs>

## Examples

<At least 2 examples using long option names>

```sh
m365 <service> <noun> <verb> --<option> <value>
```

## Response

<Tabs>
  <TabItem value="JSON">

  ```json
  { ... }
  ```

  </TabItem>
  <TabItem value="Text">

  ```text
  ...
  ```

  </TabItem>
  <TabItem value="CSV">

  ```csv
  ...
  ```

  </TabItem>
  <TabItem value="Markdown">

  ```md
  ...
  ```

  </TabItem>
</Tabs>
````

#### Rules

- Required options: angle brackets `<option>`. Optional: square brackets `[option]`.
- Examples use **long** option names, start with `m365`.
- Normalize data: tenant → `contoso`, no real PII.
- List commands: JSON wrapped in `[ ]` with one item.
- No output commands: write `The command won't return a response on success.`
- Add Remarks section between Options and Examples if needed (preview API, 0-based index, etc.).

#### Register in sidebar

Edit `docs/src/config/sidebars.ts`. Find the correct service section, locate or create the command group, add the doc entry alphabetically:

```typescript
{
  type: 'doc',
  label: '<noun> <verb>',
  id: 'cmd/<service>/<noun>/<noun>-<verb>'
}
```

### Phase 5: Verify

**STOP — Read `references/pr-checklist.md` and verify every item passes before declaring done.**

1. Run `npm run build` — must pass.
2. Run `npm test` — all tests green.
3. **STOP — Check the coverage output for the new command file.** All four metrics (Stmts, Branch, Funcs, Lines) must show 100%. If any metric is below 100%, add tests for the uncovered lines/branches and re-run until all are 100%. Do NOT proceed until this passes.
4. Walk through every checklist item in `references/pr-checklist.md`.
5. Fix any failures before proceeding.

Only after all checks pass is the command complete.
