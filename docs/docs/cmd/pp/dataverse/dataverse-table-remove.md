# pp dataverse table remove

Removes a dataverse table in a given environment

## Usage

```sh
pp dataverse table remove [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment to remove a table from.

`-n, --name<name>`
: The name of the dataverse table to remove.

`--confirm`
: Don't prompt for confirmation

`-a, --asAdmin`
: Set, to remove the dataverse table as admin for environments you are not a member of.

--8<-- "docs/cmd/_global.md"

## Examples

Removes a dataverse table in a given environment

```sh
m365 pp dataverse table remove -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "aaduser"
```

Removes a dataverse table in a given environment as Admin

```sh
m365 pp dataverse table remove -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "aaduser" --asAdmin
```

Removes a dataverse table in a given environment without prompting for confirmation

```sh
m365 pp dataverse table remove -e "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --name "aaduser" --confirm
```

## Response

The command won't return a response on success.
