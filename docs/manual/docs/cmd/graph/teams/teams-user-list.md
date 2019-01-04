# graph teams user list

Lists Microsoft Teams teams users for a specified team.

## Usage

```sh
graph teams user list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The GroupId of the team
`-r, --role <type>`|Filter the results to only users with the given role: Owner|Member|Guest
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To list  users in  Microsoft Teams, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

List all users and their role in the selected team

```sh
graph teams user list --i '00000000-0000-0000-0000-000000000000'
```

List all owners and their role in the selected team

```sh
graph teams user list --i '00000000-0000-0000-0000-000000000000' -r Owner
```

 List all guests and their role in the selected team

```sh
graph teams user list --i '00000000-0000-0000-0000-000000000000' -r Guest
```