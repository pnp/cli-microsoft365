# graph teams user add

Adds user to the specified Microsoft Teams team

## Usage

```sh
graph teams user add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the Teams team to which to add the user
`-n, --userName <userName>`|User's UPN (user principal name, eg. johndoe@example.com)
`-r, --role [role]`|The role to be assigned to the new user: `Owner|Member`. Default `Member`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To add user to the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Add a new member to the specified team

```sh
graph teams user add --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'
```

Add a new owner to the specified team

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner
```