# teams deeplink tab generate

Generates a Microsoft Teams deep link from an existing Tab in a Channel

## Usage

```sh
teams deeplink tab generate [options]
```

## Options

Option|Description
------|-----------
`--help`| output usage information
`-i, --teamId <teamId>`|The ID of the team where the tab exists
`-c, --channelId <channelId>`|The ID of the channel where the tab exists
`-t, --tabId <tabId>`|The ID of the tab to generate the deep link from
`-l, --label <label>`|The label to use in the deep link
`-m, --tabType <TabTypeOptions>`|The tab type. Allowed values `Static`, `Configurable`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--pretty`|Prettifies `json` output
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Generates a Microsoft Teams deep link from an existing Tab with id 1432c9da-8b9c-4602-9248-e0800f3e3f07


Get deeplink for tab with id, for a configurable tab
```sh
teams deeplink tab generate --teamId '00000000-0000-0000-0000-000000000000' --channelId '19:00000000000000000000000000000000@thread.skype' --tabId '1432c9da-8b9c-4602-9248-e0800f3e3f07' --label 'MyLabel' --tabType 'Configurable'
```

Get deeplink for tab with id, for a static tab
```sh
teams deeplink tab generate --teamId '00000000-0000-0000-0000-000000000000' --channelId '19:00000000000000000000000000000000@thread.skype' --tabId '1432c9da-8b9c-4602-9248-e0800f3e3f07' --label 'MyLabel' --tabType 'Static'
```