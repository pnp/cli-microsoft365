# graph o365group remove

Removes an Office 365 Group

## Usage

```sh
graph o365group remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the Office 365 Group to remove
`--confirm`|Don't prompt for confirming removing the group
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To remove an Office 365 Group, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

If the specified _id_ doesn't refer to an existing group, you will get a `Resource does not exist` error.

## Examples

Remove group with id _28beab62-7540-4db1-a23f-29a6018a3848_. Will prompt for confirmation before removing the group

```sh
graph o365group remove --id 28beab62-7540-4db1-a23f-29a6018a3848
```

Remove group with id _28beab62-7540-4db1-a23f-29a6018a3848_ without prompting for confirmation

```sh
graph o365group remove --id 28beab62-7540-4db1-a23f-29a6018a3848 --confirm
```