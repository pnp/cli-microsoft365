# teams message list

Lists all messages from a channel in a Microsoft Teams team

## Usage

```sh
m365 teams message list [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the team where the channel is located

`-c, --channelId <channelId>`
: The ID of the channel for which to list messages

`-s, --since [since]`
: Date (ISO standard, dash separator) to get delta of messages from (in last 8 months)

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on an API that is currently in preview and is subject to change once the API reached general availability.

You can only retrieve a message from a Microsoft Teams team if you are a member of that team.

## Examples

List the messages from a channel of the Microsoft Teams team

```sh
m365 teams message list --teamId fce9e580-8bba-4638-ab5c-ab40016651e3 --channelId 19:eb30973b42a847a2a1df92d91e37c76a@thread.skype
```

List the messages from a channel of the Microsoft Teams team that have been created or modified since the date specified by the `--since` parameter (WARNING: only captures the last 8 months of data)

```sh
m365 teams message list --teamId fce9e580-8bba-4638-ab5c-ab40016651e3 --channelId 19:eb30973b42a847a2a1df92d91e37c76a@thread.skype --since 2019-12-31T14:00:00Z
```
