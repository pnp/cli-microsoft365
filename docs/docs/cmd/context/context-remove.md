# context remove

Removes the CLI for Microsoft 365 context in the current working folder

## Usage

```sh
m365 context remove [options]
```

## Options

`--confirm`
: Don't prompt for confirmation to remove the context

--8<-- "docs/cmd/_global.md"

## Examples

Removes the CLI for Microsoft 365 context in the current working folder

```sh
m365 context remove
```

Removes the CLI for Microsoft 365 context in the current working folder and does not prompt for confirmation before deleting.

```sh
m365 context remove --confirm
```

## Response

The command won't return a response on success.
