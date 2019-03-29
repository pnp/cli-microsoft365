# graph teams messagingsettings set

Updates messaging settings of a Microsoft Teams team

## Usage

```sh
graph teams messagingsettings set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the Microsoft Teams team for which to update messaging settings
`--allowUserEditMessages [allowUserEditMessages]`|Set to `true` to allow users to edit messages and to `false` to disallow it
`--allowUserDeleteMessages [allowUserDeleteMessages]`|Set to `true` to allow users to delete messages and to `false` to disallow it
`--allowOwnerDeleteMessages [allowOwnerDeleteMessages]`|Set to `true` to allow owner to delete messages and to `false` to disallow it
`--allowTeamMentions [allowTeamMentions]`|Set to `true` to allow @team or @[team name] mentions and to `false` to disallow it
`--allowChannelMentions [allowChannelMentions]`|Set to `true` to allow @channel or @[channel name] mentions and to `false` to disallow it
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To update messaging settings of the specified Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Allow users to edit messages in channels

```sh
graph teams messagingsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowUserEditMessages true
```

Disallow users to delete messages in channels

```sh
graph teams messagingsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowUserDeleteMessages false
```