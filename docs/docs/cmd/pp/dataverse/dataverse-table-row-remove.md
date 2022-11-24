# pp dataverse table row remove

Removes a row from a dataverse table in a given environment.

## Usage

```sh
pp dataverse table row remove [options]
```

## Options

`-e, --environment <environment>`
: The name of the environment to remove a row from a table from.

`-i, --id <id>`
: The id of the row to remove.

`--entitySetName [entitySetName]`
: The entity set name of the table. Specify either `entitySetName` or `tableName` but not both

`--tableName [tableName]`
: The name of the table. Specify either `entitySetName` or `tableName` but not both

`--confirm`
: Don't prompt for confirmation

`--asAdmin`
: Set, to remove the row from the dataverse table as admin for environments you are not a member of.

--8<-- "docs/cmd/_global.md"

## Examples

Removes a row from a dataverse table in a given environment

```sh
m365 pp dataverse table row remove --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --tableName "aadusers" --id "21d01cf4-356c-ed11-9561-000d3a4bbea4"
```

Removes a row from a dataverse table in a given environment as Admin

```sh
m365 pp dataverse table row remove --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --tableName "aadusers" --id "21d01cf4-356c-ed11-9561-000d3a4bbea4" --asAdmin
```

Removes a row from a dataverse table in a given environment without prompting for confirmation

```sh
m365 pp dataverse table row remove --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --tableName "aadusers" --id "21d01cf4-356c-ed11-9561-000d3a4bbea4" --confirm
```

Removes a row from a dataverse table in a given environment based on the entity set name without prompting for confirmation

```sh
m365 pp dataverse table row remove --environment "Default-2ca3eaa5-140f-4175-8261-3272edf9f339" --entitySetName "cr6c3_accounts" --id "21d01cf4-356c-ed11-9561-000d3a4bbea4" --confirm
```

## Response

The command won't return a response on success.
