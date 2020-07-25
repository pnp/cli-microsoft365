# teams guestsettings set

Updates guest settings of a Microsoft Teams team

## Usage

```sh
m365 teams guestsettings set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the Teams team for which to update settings

`--allowCreateUpdateChannels [allowCreateUpdateChannels]`
: Set to `true` to allow guests to create and update channels and to `false` to disallow it

`--allowDeleteChannels [allowDeleteChannels]`
: Set to `true` to allow guests to create and update channels and to `false` to disallow it

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Allow guests to create and edit channels

```sh
m365 teams guestsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowCreateUpdateChannels true
```

Disallow guests to delete channels

```sh
m365 teams guestsettings set --teamId '00000000-0000-0000-0000-000000000000' --allowDeleteChannels false
```