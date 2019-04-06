# graph teams channel get

Gets information about the specific Microsoft Teams team channel

## Usage

```sh
graph teams channel get [options]
```

## Options

Option|Description
------|-----------
`--help`| output usage information
`-i, --teamId <teamId>`|The ID of the team to which the channel belongs
`-c, --channelId <channelId>`|The ID of the channel for which to retrieve more information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command

## Remarks

To get information of a specified channel in the given Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples
  
Get information about Microsoft Teams team channel with id _19:493665404ebd4a18adb8a980a31b4986@thread.skype_

```sh
graph teams channel get --teamId '00000000-0000-0000-0000-000000000000' --channelId '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
```