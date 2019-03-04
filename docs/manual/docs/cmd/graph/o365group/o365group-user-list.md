# graph o365group user list

Lists users for the specified Office 365 Group

## Usage

```sh
graph o365group user list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --groupId <groupId>`|The ID of the group for which to list users
`-r, --role <type>`|Filter the results to only users with the given role: `Owner|Member|Guest`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To list users in the specified Office 365 Group, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

List all users and their role in the specified Office 365 Group

```sh
graph o365group user list --groupId '00000000-0000-0000-0000-000000000000'
```

List all owners and their role in the specified Office 365 Group

```sh
graph o365group user list --groupId '00000000-0000-0000-0000-000000000000' --role Owner
```

 List all guests and their role in the specified Office 365 Group

```sh
graph o365group user list --groupId '00000000-0000-0000-0000-000000000000' --role Guest
```