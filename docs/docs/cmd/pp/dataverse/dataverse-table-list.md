# pp dataverse table list

Lists dataverse tables in a given environment

## Usage

```sh
pp dataverse table list [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment to list all tables for

`-a, --asAdmin`
: Set, to retrieve the dataverse tables as admin for environments you are not a member of.

--8<-- "docs/cmd/_global.md"

## Examples

List all tables for the given environment

```sh
m365 pp dataverse table list -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339"
```

List all tables for the given environment as Admin

```sh
m365 pp dataverse table list -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --asAdmin
```
