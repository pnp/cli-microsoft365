# teams membersettings set

Updates member settings of a Microsoft Teams team

## Usage

```sh
m365 teams membersettings set [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the Teams team for which to update settings

`--allowAddRemoveApps [allowAddRemoveApps]`
: Set to `true` to allow members to add and remove apps and to `false` to disallow it

`--allowCreateUpdateChannels [allowCreateUpdateChannels]`
: Set to `true` to allow members to create and update channels and to `false` to disallow it

`--allowCreateUpdateRemoveConnectors [allowCreateUpdateRemoveConnectors]`
: Set to `true` to allow members to create, update and remove connectors and to `false` to disallow it

`--allowCreateUpdateRemoveTabs [allowCreateUpdateRemoveTabs]`
: Set to `true` to allow members to create, update and remove tabs and to `false` to disallow it

`--allowDeleteChannels [allowDeleteChannels]`
: Set to `true` to allow members to create and update channels and to `false` to disallow it

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Allow members to create and edit channels

```sh
m365 teams membersettings set --teamId '00000000-0000-0000-0000-000000000000' --allowCreateUpdateChannels true
```

Disallow members to add and remove apps

```sh
m365 teams membersettings set --teamId '00000000-0000-0000-0000-000000000000' --allowAddRemoveApps false
```