# graph teams channel list

Lists channels in the specified Microsoft Teams team

## Usage

```sh
graph teams channel list [options]
```

## Options

Option|Description
------|-----------
`--help`| output usage information
`-i, --teamId <teamId>`|The ID of the team to list the channels of
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command

## Remarks

To list the channels in a Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples
  
List the channels in a specified Microsoft Teams team

```sh
graph teams channel list --teamId 00000000-0000-0000-0000-000000000000
```