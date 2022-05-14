# teams channel member set

Updates the role of the specified member in the specified Microsoft Teams private team channel

## Usage

```sh
m365 teams channel member set [options]
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

`-r, --role <role>`
: The role to be assigned to the user: owner, member.

--8<-- "docs/cmd/_global.md"

## Examples
  
Updates the role of the user _johndoe@example.com_ to owner in the Microsoft Teams team with id 00000000-0000-0000-0000-000000000000 and channel id 19:00000000000000000000000000000000@thread.skype

```sh
m365 teams channel member set --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --userName "johndoe@example.com" --role owner
```

Updates the role of the user with id 00000000-0000-0000-0000-000000000000 to member in the Microsoft Teams team with name _Team Name_ and channel with name _Channel Name_

```sh
m365 teams channel member set --teamName "Team Name" --channelName "Channel Name" --userId 00000000-0000-0000-0000-000000000000 --role member
```