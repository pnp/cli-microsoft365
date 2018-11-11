# graph teams channel message get

Retrieves a message from a channel in a Microsoft Teams team

## Usage

```sh
graph teams channel message get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team where the channel is located
`-c, --channelId <channelId>`|The ID of the channel that contains the message
`-m, --messageId <messageId>`|The ID of the message to retrieve
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

To retrieve a message from a Microsoft Teams channel, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

You can only retrieve a message from a Microsoft Teams team if you are a member of that team.

## Examples

Retrieve the specified message from a channel of the Microsoft Teams team

```sh
graph teams channel message get --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --messageId 1540747442203
```