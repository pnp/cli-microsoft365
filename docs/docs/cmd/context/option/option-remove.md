# context option remove

Removes an already available name from local context file.

## Usage

```sh
m365 context option remove [options]
```

## Options

`-n, --name <name>`
: The name of the option which will be deleted from the context

`--confirm`
: Don't prompt for confirming removing the option

--8<-- "docs/cmd/_global.md"

## Examples

Removes an already available name from the local context file

```sh
m365 context option remove --name "listName"
```

Removes an already available name from the local context file without confirmation

```sh
m365 context option remove --name "listName" --confirm
```

## Response

The command won't return a response on success.
