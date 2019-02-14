# graph teams user set

Promote or demote the specified member or owner for the specified Microsoft Teams team

## Usage

```sh
graph teams user set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team for which to promote or demote the specified member or owner
`-n, --userName <userName>`|User's UPN (user principal name, eg. johndoe@example.com)
`-r, --role <type>`| The role to apply to the specified user: `Owner|Member`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To promote or demote members and owners in the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Promote the specified member to Owner  

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner
```

Demote the specified member to Member  

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Member
```
