# teams message reply list

Retrieves replies to a message from a channel in a Microsoft Teams team

## Usage

```sh
m365 teams message reply list [options]
```

## Options

`-i, --teamId <teamId>`
: The ID of the team where the channel is located

`-c, --channelId <channelId>`
: The ID of the channel that contains the message

`-m, --messageId <messageId>`
: The ID of the message to retrieve replies for

--8<-- "docs/cmd/_global.md"

## Remarks

You can only retrieve replies to a message from a Microsoft Teams team if you are a member of that team.

## Examples

Retrieve the replies from a specified message from a channel of the Microsoft Teams team

```sh
m365 teams message reply list --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --messageId 1540747442203
```
