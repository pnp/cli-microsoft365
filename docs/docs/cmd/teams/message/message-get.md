# teams message get

Retrieves a message from a channel in a Microsoft Teams team

## Usage

```sh
m365 teams message get [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the team where the channel is located

`-c, --channelId <channelId>`
: The ID of the channel that contains the message

`-m, --messageId <messageId>`
: The ID of the message to retrieve

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

Retrieve the specified message from a channel of the Microsoft Teams team

```sh
m365 teams message get --teamId 5f5d7b71-1161-44d8-bcc1-3da710eb4171 --channelId 19:88f7e66a8dfe42be92db19505ae912a8@thread.skype --messageId 1540747442203
```