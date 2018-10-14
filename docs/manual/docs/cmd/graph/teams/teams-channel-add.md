# graph teams channel add

Adds channel to the Microsoft Teams team in the tenant

## Usage

```sh
graph teams channel add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --groupId`|The group id to add the channel
`-n, --name`|The name of the channel
`-d, --description`|The description of the channel
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To add a channel top Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

You can only add a channel to the Microsoft Teams team you are a member of.

## Examples

Add channel to the Microsoft Teams team in the tenant

```sh
graph teams channel add -i 6703ac8a-c49b-4fd4-8223-28f0ac3a6402 -n office365cli -d development
```