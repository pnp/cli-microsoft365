# context option remove

Removes an option by defined option from context

## Usage

```sh
m365 context option remove [options]
```

## Options

`-n, --name <name>`
: The option name which will be deleted from the context

`--confirm`
: Don't prompt for confirmation to remove the context option

--8<-- "docs/cmd/_global.md"

## Examples

Removes a CLI for Microsoft 365 context option in the current working folder

```sh
m365 context option remove --name "listName"
```

Removes a CLI for Microsoft 365 context option in the current working folder and does not prompt for confirmation before deleting.

```sh
m365 context option remove --name "listName" --confirm
```

## Response

The command won't return a response on success.
