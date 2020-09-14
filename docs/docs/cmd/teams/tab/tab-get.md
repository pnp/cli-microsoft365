# teams tab get

Gets information about the specified Microsoft Teams tab

## Usage

```sh
teams tab get [options]
```

## Options

Option|Description
------|-----------
`--help`| output usage information
`-i, --teamId <teamId>`|The ID of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both
`--teamName [teamName]`|The display name of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both
`-c, --channelId <channelId>`|The ID of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both
`--channelName [channelName]`|The display name of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both
`-t, --tabId [tabId]`|The ID of the Microsoft Teams tab. Specify either tabId or tabName but not both
`--tabName [tabName]`|The display name of the Microsoft Teams tab. Specify either tabId or tabName but not both
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `json,text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

You can only retrieve tabs for teams of which you are a member.

## Examples
  
Get url of a Microsoft Teams Tab with id 1432c9da-8b9c-4602-9248-e0800f3e3f07

```sh
teams tab get --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --tabId 1432c9da-8b9c-4602-9248-e0800f3e3f07
```

Get url of a Microsoft Teams Tab with name "Tab Name"

```sh
teams tab list --teamName "Team Name" --channelName "Channel Name" --tabName "Tab Name"
```