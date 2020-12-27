# teams messagingsettings set

Updates messaging settings of a Microsoft Teams team

## Usage

```sh
m365 teams messagingsettings set [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the Microsoft Teams team for which to update messaging settings

`--allowUserEditMessages [allowUserEditMessages]`
: Set to `true` to allow users to edit messages and to `false` to disallow it

`--allowUserDeleteMessages [allowUserDeleteMessages]`
: Set to `true` to allow users to delete messages and to `false` to disallow it

`--allowOwnerDeleteMessages [allowOwnerDeleteMessages]`
: Set to `true` to allow owner to delete messages and to `false` to disallow it

`--allowTeamMentions [allowTeamMentions]`
: Set to `true` to allow @team or @[team name] mentions and to `false` to disallow it

`--allowChannelMentions [allowChannelMentions]`
: Set to `true` to allow @channel or @[channel name] mentions and to `false` to disallow it

--8<-- "docs/cmd/_global.md"

## Examples

Allow users to edit messages in channels

```sh
m365 teams messagingsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowUserEditMessages true
```

Disallow users to delete messages in channels

```sh
m365 teams messagingsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowUserDeleteMessages false
```