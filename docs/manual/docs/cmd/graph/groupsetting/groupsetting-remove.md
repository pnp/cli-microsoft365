# graph groupsetting remove

Removes the particular group setting

## Usage

```sh
graph groupsetting remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|The ID of the group setting to remove
`--confirm`|Don't prompt for confirming removing the group setting
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To remove a group setting, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

If the specified _id_ doesn't refer to an existing group setting, you will get a `Resource does not exist` error.

## Examples

Remove group setting with id _28beab62-7540-4db1-a23f-29a6018a3848_. Will prompt for confirmation before removing the group setting

```sh
graph groupsetting remove --id 28beab62-7540-4db1-a23f-29a6018a3848
```

Remove group setting with id _28beab62-7540-4db1-a23f-29a6018a3848_ without prompting for confirmation

```sh
graph groupsetting remove --id 28beab62-7540-4db1-a23f-29a6018a3848 --confirm
```