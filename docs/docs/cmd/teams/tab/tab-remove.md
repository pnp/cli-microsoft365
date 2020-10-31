# teams tab remove

Removes a tab from the specified channel

## Usage

```sh
m365 teams tab remove [options]
```

## Options

`-h, --help`
: output usage information

`-i, --teamId <teamId>`
: The ID of the team where the tab exists

`-c, --channelId <channelId>`
: The ID of the channel to remove the tab from

`-t, --tabId <tabId>`
: The ID of the tab to remove

`--confirm`
: Don't prompt for confirmation

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json|text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Removes a tab from the specified channel. Will prompt for confirmation

```sh
m365 teams tab remove --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --tabId 06805b9e-77e3-4b93-ac81-525eb87513b8
```

Removes a tab from the specified channel without prompting for confirmation

```sh
m365 teams tab remove --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --tabId 06805b9e-77e3-4b93-ac81-525eb87513b8 --confirm
```

## Additional information

- Delete tab from channel: [https://docs.microsoft.com/en-us/graph/api/teamstab-delete?view=graph-rest-1.0](https://docs.microsoft.com/en-us/graph/api/teamstab-delete?view=graph-rest-1.0)
