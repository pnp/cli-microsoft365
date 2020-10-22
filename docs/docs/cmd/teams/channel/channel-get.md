# teams channel get

Gets information about the specific Microsoft Teams team channel

## Usage

```sh
m365 teams channel get [options]
```

## Options

`-i, --teamId [teamId]`
: The ID of the team to which the channel belongs to. Specify either teamId or teamName but not both

`--teamName [teamName]`
: The display name of the team to which the channel belongs to. Specify either teamId or teamName but not both

`-c, --channelId <channelId>`
: The ID of the channel for which to retrieve more information. Specify either channelId or channelName but not both

`--channelName [channelName]`
: The display name of the channel for which to retrieve more information. Specify either channelId or channelName but not both

--8<-- "docs/cmd/_global.md"

## Examples
  
Get information about Microsoft Teams team channel with id _19:493665404ebd4a18adb8a980a31b4986@thread.skype_

```sh
m365 teams channel get --teamId 00000000-0000-0000-0000-000000000000 --channelId '19:493665404ebd4a18adb8a980a31b4986@thread.skype'
```

Get information about Microsoft Teams team channel with name _Channel Name_

```sh
m365 teams channel get --teamName "Team Name" --channelName "Channel Name"
```