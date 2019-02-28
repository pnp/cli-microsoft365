# graph teams user set

Updates role of the specified user in the given Microsoft Teams team

## Usage

```sh
graph teams user set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team for which to change the user's role
`-n, --userName <userName>`|UPN of the user for whom to update the role (eg. johndoe@example.com)
`-r, --role <role>`|Role to set for the given user in the specified team. Allowed values: `Owner|Member`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To update role of the given user in the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

The command will return an error if the user already has the specified role in the given Microsoft Teams team.

## Examples

Promote the specified user to owner of the given Microsoft Teams team

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner
```

Demote the specified user from owner to member in the given Microsoft Teams team

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Member
```