# graph teams remove

Removes the specified Microsoft Teams team

## Usage

```sh
graph teams remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the Teams team to remove
`--confirm`|Don't prompt for confirming removing the specified team
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To remove the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

When deleted, Office 365 groups are moved to a temporary container and can be restored within 30 days. 
After that time, they are permanently deleted. This applies only to Office 365 groups.

## Examples

Removes the specified team

```sh
graph teams remove --teamId '00000000-0000-0000-0000-000000000000'
```

Removes the specified team without confirmation

```sh
graph teams remove --teamId '00000000-0000-0000-0000-000000000000' --confirm
```

## More information

- directory resource type (deleted items): [https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0)