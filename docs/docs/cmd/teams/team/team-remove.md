# teams team remove

Removes the specified Microsoft Teams team

## Usage

```sh
m365 teams team remove [options]
```

## Options

`-i, --id [id]`
: The ID of the Microsoft Teams team to remove. Specify either id or name but not both

`-n, --name [name]`
: The display name of the Microsoft Teams team to remove. Specify either id or name but not both

`--teamId [teamId]`
: (deprecated. Use `id` instead) The ID of the Teams team to remove

`--confirm`
: Don't prompt for confirming removing the specified team

--8<-- "docs/cmd/_global.md"

## Remarks

When deleted, Microsoft 365 groups are moved to a temporary container and can be restored within 30 days. After that time, they are permanently deleted. This applies only to Microsoft 365 groups.

If the command finds multiple Microsoft Teams teams with the specified name, it will prompt you to disambiguate which team it should use, listing the discovered team IDs.

## Examples

Removes the specified Microsoft Teams team with id _00000000-0000-0000-0000-000000000000_

```sh
m365 teams team remove --id 00000000-0000-0000-0000-000000000000
```

Removes the specified Microsoft Teams team with name _Team Name_

```sh
m365 teams team remove --name "Team Name"
```

Removes the specified team without confirmation

```sh
m365 teams team remove --id 00000000-0000-0000-0000-000000000000 --confirm
```

## More information

- directory resource type (deleted items): [https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0)
