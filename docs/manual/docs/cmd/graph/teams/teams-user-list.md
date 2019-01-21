# graph teams user list

Lists users for the specified Microsoft Teams team

## Usage

```sh
graph teams user list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team for which to list users
`-r, --role <type>`|Filter the results to only users with the given role: `Owner|Member|Guest`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To list users in the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

List all users and their role in the specified team

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000'
```

List all owners and their role in the specified team

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000' --role Owner
```

 List all guests and their role in the specified team

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000' --role Guest
```