# pp card list

Lists Microsoft Power Platform cards in the specified Power Platform environment.

## Usage

```sh
pp card list [options]
```

## Options

`-e, --environment <environment>`
The name of the environment for which to retrieve cards.

`-a, --asAdmin`
Run the command as admin and retrieve cards for environments you do not have explicitly assigned permissions to.

--8<-- "docs/cmd/_global.md"

## Examples

List cards in the environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_.

```sh
m365 pp card list --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5"
```

List cards in the environment named _Default-d87a7535-dd31-4437-bfe1-95340acd55c5_ as admin.

```sh
m365 pp card list --environment "Default-d87a7535-dd31-4437-bfe1-95340acd55c5" --asAdmin
```
