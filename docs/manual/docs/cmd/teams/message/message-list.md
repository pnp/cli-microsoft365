# teams message list

Lists all messages from a channel in a Microsoft Teams team

## Usage

```sh
teams message list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team where the channel is located
`-c, --channelId <channelId>`|The ID of the channel for which to list messages
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

You can only retrieve a message from a Microsoft Teams team if you are a member of that team.

## Examples

Retrieve the specified message from a channel of the Microsoft Teams team

```sh
teams message list --teamId fce9e580-8bba-4638-ab5c-ab40016651e3 --channelId 19:eb30973b42a847a2a1df92d91e37c76a@thread.skype
```