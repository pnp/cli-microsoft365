# teams conversationmember add

Adds a conversation member in a private channel

## Usage

```sh
m365 teams conversationmember add [options]
```

## Options

`-i, --teamId [teamId]`
: The ID of the team where the channel is located. Specify either `teamId` or `teamName`, but not both.

`--teamName [teamName]`
: The name of the team where the channel is located. Specify either `teamId` or `teamName`, but not both.

`-c, --channelId [channelId]`
: The private channel's ID. Specify either `channelId` or `channelName`, but not both.

`--channelName [channelName]`
: The private channel's name. Specify either `channelId` or `channelName`, but not both.

`--userId [userId]`
: The user's ID or principal name. You can also pass a comma separated list of userIds.

`--userDisplayName [userDisplayName]`
: The display name of a user. You can also pass a comma separated list of display names.

`--owner`
: Assign the user the owner role. Defaults to member permissions.

--8<-- "docs/cmd/_global.md"

## Remarks

At least one owner must be assigned to a private channel.

You can only add members and owners of a Team to a private channel.

## Examples

Add members to a channel based on their id or user principal name

```sh
m365 teams conversationmember add --teamId 47d6625d-a540-4b59-a4ab-19b787e40593 --channelId 19:586a8b9e36c4479bbbd378e439a96df2@thread.skype --userId "85a50aa1-e5b8-48ac-b8ce-8e338033c366,john.doe@contoso.com"
```

Add owners to a channel based on their display names

```sh
m365 teams conversationmember add --teamName "Human Resources" --channelName "Private Channel" --userDisplayName "Anne Matthews,John Doe" --owner
```
