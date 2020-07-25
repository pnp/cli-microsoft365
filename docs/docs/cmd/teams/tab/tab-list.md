# teams tab list

Lists tabs in the specified Microsoft Teams channel

## Usage

```sh
m365 teams tab list [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the Microsoft Teams team where the channel is located

`-c, --channelId <channelId>`
: The ID of the channel for which to list tabs

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

You can only retrieve tabs for teams of which you are a member.

Tabs _Conversations_ and _Files_ are present in every team and therefore not included in the list of available tabs.

## Examples
  
List all tabs in a Microsoft Teams channel

```sh
m365 teams tab list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype
```

Include all the values from the tab configuration and associated teams app

```sh
m365 teams tab list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --output json
```