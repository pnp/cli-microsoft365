# graph o365group remove

Removes the specified Office 365 Group

## Usage

```sh
graph o365group remove --id 
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the Office 365 Group to remove
`--confirm`|Don't prompt for confirming removing the group
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

To remove a Office 365 Group, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

## Examples

Remove the Office 365 Group with id _28beab62-7540-4db1-a23f-29a6018a3848_ and prompt for confirmation before removing the group.

```sh
graph o365group remove --id 28beab62-7540-4db1-a23f-29a6018a3848
```

Remove the Office 365 Group with id _28beab62-7540-4db1-a23f-29a6018a3848_ without prompting
    for confirmation
  
```sh
graph o365group remove --id 28beab62-7540-4db1-a23f-29a6018a3848 --confirm
```