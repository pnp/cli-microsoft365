---
name: zod-migration
description: 'Migrate CLI for Microsoft 365 commands from legacy initOptions/initValidators/initTelemetry pattern to Zod schema validation. Use when asked to "migrate to zod", "upgrade command to zod", "convert command validation to zod", or when working on commands that still use the old pattern.'
---

# Zod Migration

Migrate a CLI for Microsoft 365 command from the legacy `#initOptions()`/`#initValidators()`/`#initTelemetry()` pattern to Zod schema-based validation.

## When to Use

- Command still has `#initOptions()`, `#initValidators()`, or `#initTelemetry()` methods
- Command does not have a `schema` getter or exported `options` Zod schema

## Pre-flight

1. Read the command source file (`.ts`) and its test file (`.spec.ts`)
2. Identify: options, aliases, validators, telemetry, option sets, types
3. Check if the command has `autocomplete` values on any options
4. Check whether any command option representing a GUID or UPN accepts the runtime tokens `@meid` or `@meusername`

## Procedure

### Step 1: Define the Zod Schema (command `.ts` file)

Replace imports and add schema definition **above** the class:

```typescript
import { z } from 'zod';
import { globalOptionsZod } from '../../../../Command.js';
// Remove: import GlobalOptions from '../../../../GlobalOptions.js';

export const options = z.strictObject({
  ...globalOptionsZod.shape,
  // Add command-specific options here
});

declare type Options = z.infer<typeof options>;

interface CommandArgs {
  options: Options;
}
```

#### Option Type Mapping

| Old Pattern | Zod Equivalent |
|---|---|
| `option: '--name <name>'` (required string) | `name: z.string()` |
| `option: '--name [name]'` (optional string) | `name: z.string().optional()` |
| `option: '-n, --name <name>'` (with alias) | `name: z.string().alias('n')` |
| `option: '--force'` (boolean flag) | `force: z.boolean().optional()` |
| `option: '--count <count>'` (in types.string array) | `count: z.string()` (keep as string, convert in commandAction if needed) |
| Option with `autocomplete: [...]` | `z.enum([...])` (preferred) or `z.string()` with `.refine()` |

#### Critical Rules for Schema Definition

1. **ALWAYS use `z.strictObject()`** — not `z.object()`. Strict rejects unknown options.
2. **ALWAYS spread `...globalOptionsZod.shape`** — includes debug, verbose, output, query.
3. **ALWAYS export `options`** — the spec file imports it for typing.
4. **NEVER use `z.uuid()` for fields that accept `@meid` token** — use `.refine()` with custom GUID validation instead (see below).
5. **Preserve case-insensitive behavior** — if old validator lowercased input before comparing, add `.transform(v => v.toLowerCase())` or use `.refine()` with case-insensitive check.
6. **Keep option names identical** — do not rename options during migration (breaking change).
7. **Use `z.enum()` for options with fixed `autocomplete` values** — this preserves shell completion metadata.

#### GUID Fields That Accept `@meid`

```typescript
// WRONG - rejects @meid token
id: z.string().uuid()

// CORRECT - allows @meid token
id: z.string().refine(val => validation.isValidGuid(val), {
  message: 'The value must be a valid GUID.'
})
```

#### Comma-Separated String Fields

When an option accepts comma-separated values (e.g., `--scopes Sites.Read.All,Sites.ReadWrite.All`), use `.transform()` to split into an array if the command processes them as arrays:

```typescript
scopes: z.string().transform(value => value.split(',').map(s => s.trim())).alias('s')
```

#### Numeric Values Received as Strings

CLI arguments are strings. If the command previously relied on yargs numeric coercion (via `types.string` exclusion), handle conversion explicitly in `commandAction`:

```typescript
// In commandAction:
const pageSize = Number(args.options.value);
```

### Step 2: Add Schema Getter to the Class

```typescript
public get schema(): z.ZodType | undefined {
  return options;
}
```

### Step 3: Remove Legacy Methods

Delete these methods entirely from the class:
- `constructor()` (if it only called `super()` + init methods)
- `#initOptions()`
- `#initValidators()`
- `#initTelemetry()`
- `#initTypes()`
- `#initOptionSets()`

### Step 4: Add `getRefinedSchema()` for Cross-Field Validation

Use `getRefinedSchema()` when:
- Command has option sets (mutually exclusive or "one of" requirements)
- Command has conditional validation (option B required when option A is set)
- Command had validators with cross-field logic

#### Option Set Pattern (exactly one of N options required)

```typescript
public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
  return schema
    .refine(opts => [opts.id, opts.name].filter(x => x !== undefined).length === 1, {
      message: `Specify either 'id' or 'name', but not both.`,
      params: {
        customCode: 'optionSet',
        options: ['id', 'name']
      }
    });
}
```

#### Required Dependency Pattern (B required when A is provided)

```typescript
public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
  return schema
    .refine(opts => !opts.cardData || opts.card, {
      error: 'When you specify cardData, you must also specify card.',
      path: ['cardData'],
      params: {
        customCode: 'required'
      }
    });
}
```

#### At-Least-One-Update Pattern (set commands)

```typescript
public getRefinedSchema(schema: typeof options): z.ZodObject<any> | undefined {
  return schema
    .refine(opts => opts.description || opts.status || opts.owner, {
      message: 'Specify at least one property to update.',
      params: {
        customCode: 'required'
      }
    });
}
```

### Step 5: Handle Validators That Must Stay in Schema

**CRITICAL:** When a command defines `schema`, the CLI does NOT run `this.validators`. All validation MUST live in the schema or `getRefinedSchema()`.

Move file-system checks, URL validation, and other validators into `.refine()` calls:

```typescript
export const options = z.strictObject({
  ...globalOptionsZod.shape,
  filePath: z.string()
    .refine(val => fs.existsSync(val), {
      message: 'Specified file does not exist.'
    })
});
```

### Step 6: Handle Loose Schemas (commands accepting unknown options)

Some commands (like `user-set`, `user-add`, `groupsetting-set`) accept arbitrary options passed through to the API. Use `z.object()` instead of `z.strictObject()`:

```typescript
export const options = z.object({
  ...globalOptionsZod.shape,
  id: z.string()
}).catchall(z.unknown());
```

### Step 7: Migrate the Test File (`.spec.ts`)

#### 7a. Update imports

```typescript
// Old:
import command from './command-name.js';
// New:
import command, { options } from './command-name.js';
```

#### 7b. Add schema variable declaration

```typescript
describe(commands.COMMAND_NAME, () => {
  let commandInfo: CommandInfo;
  let commandOptionsSchema: typeof options;  // ADD THIS

  before(() => {
    commandInfo = cli.getCommandInfo(command);
    commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options;  // ADD THIS
    // ...
  });
```

#### 7c. Convert validation tests to use `safeParse`

```typescript
// Old:
it('fails validation if id is not valid', async () => {
  const actual = await command.validate({ options: { id: 'invalid' } }, commandInfo);
  assert.notStrictEqual(actual, true);
});

// New:
it('fails validation if id is not valid', () => {
  const actual = commandOptionsSchema.safeParse({ id: 'invalid' });
  assert.strictEqual(actual.success, false);
});
```

Note: Validation tests become **synchronous** (no `async`).

#### 7d. Convert action tests to use `commandOptionsSchema.parse()`

```typescript
// Old:
await command.action(logger, { options: { id: 'abc', verbose: true } } as any);

// New:
await command.action(logger, { options: commandOptionsSchema.parse({ id: 'abc', verbose: true }) });
```

**CRITICAL:** Use `.parse()` (throws) for action tests, `.safeParse()` for validation tests.

**Exception for `@meid`/`@meusername` tokens:** Tests that use runtime tokens like `@meid` or `@meusername` (which are replaced by `loadValuesFromAccessToken` before schema validation) must keep `as any` since these values intentionally bypass Zod:

```typescript
// @meid is a runtime token - cannot go through parse
await command.action(logger, { options: { id: '@meid' } as any });
```

#### 7e. Remove legacy "supports specifying" tests

Delete tests like:
```typescript
// DELETE these — schema handles option registration
it('supports specifying id', () => {
  const options = command.options;
  let containsOption = false;
  options.forEach(o => { ... });
  assert(containsOption);
});
```

#### 7f. Add required validation tests

**ALWAYS add these tests:**

For commands that reject unknown options:
```typescript
it('fails validation with unknown options', () => {
  const actual = commandOptionsSchema.safeParse({
    id: 'valid-id',
    unknownOption: 'value'
  });
  assert.strictEqual(actual.success, false);
});
```

For commands where all specific options are optional:
```typescript
it('passes validation with no options', () => {
  const actual = commandOptionsSchema.safeParse({});
  assert.strictEqual(actual.success, true);
});
```

## Checklist

Before submitting, verify:

- [ ] `options` is **exported** from the command file
- [ ] Schema uses `z.strictObject()` (unless command accepts unknown options)
- [ ] Schema spreads `...globalOptionsZod.shape`
- [ ] `schema` getter returns `options`
- [ ] All old init methods (`#initOptions`, `#initValidators`, `#initTelemetry`, `#initTypes`, `#initOptionSets`) are removed
- [ ] Constructor removed (if it only called init methods)
- [ ] No `this.validators` or `this.options.unshift(...)` remain
- [ ] All validators moved to schema `.refine()` or `getRefinedSchema()`
- [ ] `getRefinedSchema()` uses `params: { customCode: 'optionSet', options: [...] }` for option sets
- [ ] `getRefinedSchema()` uses `params: { customCode: 'required' }` for conditional requirements
- [ ] GUID fields that accept `@meid` do NOT use `z.uuid()`
- [ ] Options with `autocomplete` values use `z.enum()` when possible
- [ ] Case-insensitive validation preserved where it existed before
- [ ] No option renames (no breaking changes)
- [ ] Spec file imports `{ options }` from the command file
- [ ] Spec declares `commandOptionsSchema: typeof options`
- [ ] Spec initializes `commandOptionsSchema = commandInfo.command.getSchemaToParse() as typeof options`
- [ ] All action tests use `commandOptionsSchema.parse({...})`
- [ ] All validation tests use `commandOptionsSchema.safeParse({...})`
- [ ] Validation tests are synchronous (no `async`)
- [ ] Legacy "supports specifying" tests removed
- [ ] Test added: `'fails validation with unknown options'`
- [ ] Test added (when applicable): `'passes validation with no options'`
- [ ] No test uses `command.validate(...)` — all validation goes through schema
