# teams channel member remove

Remove the specified member from the specified Microsoft Teams private team channel

## Usage

```sh
m365 teams channel member remove [options]
```

## Alias

```sh
m365 teams conversationmember remove [options]
```

## Options

`--teamId [teamId]`
: The Id of the Microsoft Teams team. Specify either `teamId` or `teamName` but not both

`--teamName [teamName]`
: The display name of the Microsoft Teams team. Specify either `teamId` or `teamName` but not both

`--channelId [channelId]`
: The Id of the Microsoft Teams team channel. Specify either `channelId` or `channelName` but not both

`--channelName [channelName]`
: The display name of the Microsoft Teams team channel. Specify either `channelId` or `channelName` but not both

`--userName [userName]`
: User's UPN (user principal name, e.g. johndoe@example.com). Specify either userName, userId or id but not multiple.

`--userId [userId]`
: User's Azure AD Id. Specify either userName, userId or id but not multiple.

`--id [id]`
: Channel member Id of a user. Specify either userName, userId or id but not multiple.

`--confirm`
: Don't prompt for confirmation

--8<-- "docs/cmd/_global.md"

## Examples
  
Remove the user _johndoe@example.com_ from the Microsoft Teams team with id 00000000-0000-0000-0000-000000000000 and channel id 00:00000000000000000000000000000000@thread.skype

```sh
m365 teams channel member remove --teamId 00000000-0000-0000-0000-000000000000 --channelId 00:00000000000000000000000000000000@thread.skype --userName "johndoe@example.com"
```

Remove the user with id 00000000-0000-0000-0000-000000000000 from the Microsoft Teams team with name _Team Name_ and channel with name _Channel Name_

```sh
m365 teams channel member remove --teamName "Team Name" --channelName "Channel Name" --userId 00000000-0000-0000-0000-000000000000
```
