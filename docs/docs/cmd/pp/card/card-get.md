# pp card get

Gets a specific Microsoft Power Platform card in the specified Power Platform environment.

## Usage

```sh
pp card get [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment.

`-i, --id <id>`
: The id of the card.

`-a, --asAdmin`
: Run the command as admin for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

Get a specific card in a specific environment.

```sh
m365 pp card get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "408e3f42-4c9e-4c93-8aaf-3cbdea9179aa"
```

Get a specific card in a specific environment as admin.

```sh
m365 pp card get --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --id "408e3f42-4c9e-4c93-8aaf-3cbdea9179aa" --asAdmin
