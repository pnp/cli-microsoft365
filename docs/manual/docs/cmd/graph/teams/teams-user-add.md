# graph teams user add

Adds an owner or member to the specified Microsoft Teams team

## Usage

```sh
graph teams user add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team for which to list users
`-n, --userName <userName>`|User\'s UPN (user principal name - e.g. johndoe@example.com)
`-r, --role <type>`|Filter the results to only users with the given role: `Owner|Member`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To add users in the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Add a new member to the specified team

```sh
graph teams user add --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com'
```

Add a new owner to the specified team 

```sh
graph teams user list --teamId '00000000-0000-0000-0000-000000000000' --userName 'anne.matthews@contoso.onmicrosoft.com' --role Owner
```