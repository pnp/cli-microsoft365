# graph o365group restore

Restores a deleted Office 365 Group

## Usage

```sh
graph o365group restore [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the Office 365 Group to restore
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to the Microsoft Graph, using the [graph connect](../connect.md) command.

## Remarks

!!! attention
    This command is based on a Microsoft Graph API that is currently in preview and is subject to change once the API reached general availability.

To restore a deleted Office 365 Group, you have to first connect to the Microsoft Graph using the [graph connect](../connect.md) command, eg. `graph connect`.

## Examples

Restores the Office 365 Group with id _28beab62-7540-4db1-a23f-29a6018a3848_

```sh
graph o365group restore --id 28beab62-7540-4db1-a23f-29a6018a3848
```