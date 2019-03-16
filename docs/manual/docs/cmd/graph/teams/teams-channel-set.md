# graph teams channel set

Updates properties of a specified channel in the given Microsoft Teams team

## Usage

```sh
graph teams channel set [options]
```

## Options

Option|Description
------|-----------
`--help`| output usage information
`-i, --teamId <teamId>`|The ID of the team for which to update channel 
`--channelName <channelName>`|The name of the channel that needs to be updated
`--newChannelName <newChannelName>`|The new name of the channel
`--description <description>`|The description of the channel
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

To update properties of a specified channel in the given Microsoft Teams team,
, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples
  
Update properties of a specified channel in the given Microsoft Teams team with description 

```sh
o365$ graph teams channel set --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects --description "Channel for new projects"
```

Update properties of a specified channel in the given Microsoft Teams team without description 

```sh
o365$ graph teams channel set --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects
```    