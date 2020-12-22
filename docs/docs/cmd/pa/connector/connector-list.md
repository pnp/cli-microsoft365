# pa connector list

Lists custom connectors in the given environment

## Usage

```sh
m365 pa connector list [options]
```

## Alias

```sh
m365 flow connector list
```

## Options

`-e, --environment <environment>`
: The name of the environment for which to retrieve custom connectors

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

## Examples

List all custom connectors in the given environment

```sh
m365 pa connector list --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5
```
