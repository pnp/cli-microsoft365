# teams channel member list

Lists members of the specified Microsoft Teams team channel

## Usage

```sh
m365 teams channel member list [options]
```

## Alias

```sh
m365 teams conversationmember list [options]
```

## Options

`-i, --teamId [teamId]`
: The Id of the Microsoft Teams team. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the Microsoft Teams team. Specify either `teamId` or `teamName` but not both

`-c, --channelId [channelId]`
: The Id of the Microsoft Teams team channel. Specify either `channelId` or `channelName` but not both

`--channelName [channelName]`
: The display name of the Microsoft Teams team channel. Specify either `channelId` or `channelName` but not both

`-r, --role [role]`
: Filter the results to only users with the given role: owner, member, guest

--8<-- "docs/cmd/_global.md"

## Examples
  
List the members of a specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000 and channel id 19:00000000000000000000000000000000@thread.skype

```sh
m365 teams channel member list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype
```

List the members of a specified Microsoft Teams team with name _Team Name_ and channel with name _Channel Name_

```sh
m365 teams channel member list --teamName "Team Name" --channelName "Channel Name"
```

List all owners of the specified Microsoft Teams team with id 00000000-0000-0000-0000-000000000000 and channel id 19:00000000000000000000000000000000@thread.skype

```sh
m365 teams channel member list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --role owner
```
