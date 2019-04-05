# graph teams user remove

Removes the specified user from the specified Microsoft Teams team

## Usage

```sh
graph teams user remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the Teams team from which to remove the user
`-n, --userName <userName>`|User's UPN (user principal name), eg. `johndoe@example.com`
`--confirm`|Don't prompt for confirming removing the user from the specified team
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To remove user from the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

You can remove users from a Microsoft Teams team if you are a owner of that team.

## Examples

Removes user from the specified team

```sh
graph teams user remove --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'
```

Removes user from the specified team without confirmation

```sh
graph teams user remove --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --confirm
```