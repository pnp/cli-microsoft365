# graph teams channel add

Adds a channel to the specified Microsoft Teams team

## Usage

```sh
graph teams channel add [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team to add the channel to
`-n, --name <name>`|The name of the channel to add
`-d, --description [description]`|The description of the channel to add
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

To add a channel top Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

You can only add a channel to the Microsoft Teams team you are a member of.

## Examples

Add channel to the specified Microsoft Teams team

```sh
graph teams channel add --teamId 6703ac8a-c49b-4fd4-8223-28f0ac3a6402 --name office365cli --description development
```