# graph teams set

Updates settings of a Microsoft Teams team

## Usage

```sh
graph teams set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the Microsoft Teams team for which to update settings
`--displayName [displayName]`|The display name for the Microsoft Teams team
`--description [description]`|The description for the Microsoft Teams team
`--mailNickName [mailNickName]`|The mail alias for the Microsoft Teams team
`--classification [classification]`|The classification for the Microsoft Teams team
`--visibility [visibility]`|The visibility of the Microsoft Teams team. Valid values `Private|Public`
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

To update the settings of the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Set Microsoft Teams team visibility as Private

```sh
graph teams set --teamId '00000000-0000-0000-0000-000000000000' --visibility Private
```

Set Microsoft Teams team classification as MBI

```sh
graph teams set --teamId '00000000-0000-0000-0000-000000000000' --classification MBI
```