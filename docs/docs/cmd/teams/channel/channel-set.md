# teams channel set

Updates properties of the specified channel in the given Microsoft Teams team

## Usage

```sh
m365 teams channel set [options]
```

## Options

`-i, --teamId [teamId]`
: The ID of the team where the channel to update is located. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the team where the channel to update is located. Specify either `teamId` or `teamName` but not both

`-c, --channelId [channelId]`
: The ID of the channel to update. Specify either `channelId` or `channelName` but not both

`-n, --channelName [channelName]`
: The name of the channel to update. Specify either `channelId` or `channelName` but not both

`--newChannelName [newChannelName]`
: The new name of the channel

`--description [description]`
: The description of the channel

--8<-- "docs/cmd/_global.md"

## Examples
  
Set new description and display name for the channel with id

```sh
m365 teams channel set --teamId "00000000-0000-0000-0000-000000000000" --channelId 19:f3dcbb1674574677abcae89cb626f1e6@thread.skype --newChannelName Projects --description "Channel for new projects"
```

Set new display name for the channel with name

```sh
m365 teams channel set --teamName "Team Name" --channelName Reviews --newChannelName Projects
```

