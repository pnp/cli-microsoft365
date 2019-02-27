# graph teams membersettings set

Updates member settings of a Microsoft Teams team

## Usage

```sh
graph teams membersettings set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the Teams team for which to update settings
`--allowAddRemoveApps [allowAddRemoveApps]`|Set to `true` to allow members to add and remove apps and to `false` to disallow it
`--allowCreateUpdateChannels [allowCreateUpdateChannels]`|Set to `true` to allow members to create and update channels and to `false` to disallow it
`--allowCreateUpdateRemoveConnectors [allowCreateUpdateRemoveConnectors]`|Set to `true` to allow members to create, update and remove connectors and to `false` to disallow it
`--allowCreateUpdateRemoveTabs [allowCreateUpdateRemoveTabs]`|Set to `true` to allow members to create, update and remove tabs and to `false` to disallow it
`--allowDeleteChannels [allowDeleteChannels]`|Set to `true` to allow members to create and update channels and to `false` to disallow it
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To update member settings of the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Allow members to create and edit channels

```sh
graph teams membersettings set --teamId '00000000-0000-0000-0000-000000000000' --allowCreateUpdateChannels true
```

Disallow members to add and remove apps

```sh
graph teams membersettings set --teamId '00000000-0000-0000-0000-000000000000' --allowAddRemoveApps false
```