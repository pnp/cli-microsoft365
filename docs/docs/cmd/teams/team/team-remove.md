# teams team remove

Removes the specified Microsoft Teams team

## Usage

```sh
m365 teams team remove [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the Teams team to remove

`--confirm`
: Don't prompt for confirming removing the specified team

--8<-- "docs/cmd/_global.md"

## Remarks

When deleted, Microsoft 365 groups are moved to a temporary container and can be restored within 30 days. After that time, they are permanently deleted. This applies only to Microsoft 365 groups.

## Examples

Removes the specified team

```sh
m365 teams team remove --teamId '00000000-0000-0000-0000-000000000000'
```

Removes the specified team without confirmation

```sh
m365 teams team remove --teamId '00000000-0000-0000-0000-000000000000' --confirm
```

## More information

- directory resource type (deleted items): [https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0)
