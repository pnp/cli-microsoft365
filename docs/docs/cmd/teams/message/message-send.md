# teams message send

Sends a message to a channel in a Microsoft Teams team

## Usage

```sh
m365 teams message send [options]
```

## Options

`-t, --teamId <teamId>`
: The ID of the team where the channel is located.

`-c, --channelId <channelId>`
: The ID of the channel.

`-m, --message <message>`
: The message to send.

--8<-- "docs/cmd/_global.md"

## Remarks

You can only send a message to a channel in a Microsoft Teams team if you are a member of that team or channel.

## Examples

Send a message to a specified channel in a Microsoft Teams team.

```sh
m365 teams message send --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --message "Hello World"
```

Send a html formatted message to a specified channel in a Microsoft Teams team.

```sh
m365 teams message send --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --message "Hello <b>World</b>"
```

## Response

The command won't return a response on success.
