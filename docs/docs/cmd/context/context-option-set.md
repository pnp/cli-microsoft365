# context option set

Allows to add a new key to the context with value

## Usage

```sh
m365 context option set [options]
```

## Options

`-n, --name <name>`
: The option name for which to define the value

`-v, --value <value>`
: Default value for the option

--8<-- "docs/cmd/_global.md"

## Examples

Adds a new key to the CLI for Microsoft 365 context in the current working folder

```sh
m365 context option set --name 'listName' --value 'testList'
```

## Response

The command won't return a response on success.
