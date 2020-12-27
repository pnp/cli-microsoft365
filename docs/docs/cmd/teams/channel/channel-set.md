# teams channel set

Updates properties of the specified channel in the given Microsoft Teams team

## Usage

```sh
m365 teams channel set [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the team where the channel to update is located

`--channelName <channelName>`
: The name of the channel to update

`--newChannelName [newChannelName]`
: The new name of the channel

`--description [description]`
: The description of the channel

--8<-- "docs/cmd/_global.md"

## Examples
  
Set new description and display name for the specified channel in the given Microsoft Teams team

```sh
m365 teams channel set --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects --description "Channel for new projects"
```

Set new display name for the specified channel in the given Microsoft Teams team

```sh
m365 teams channel set --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects
```