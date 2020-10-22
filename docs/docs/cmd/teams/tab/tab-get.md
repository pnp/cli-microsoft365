# teams tab get

Gets information about the specified Microsoft Teams tab

## Usage

```sh
m365 teams tab get [options]
```

## Options

`--teamId [teamId]`
: The ID of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both

`--teamName [teamName]`
: The display name of the Microsoft Teams team where the tab is located. Specify either teamId or teamName but not both

`--channelId [channelId]`
: The ID of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both

`--channelName [channelName]`
: The display name of the Microsoft Teams channel where the tab is located. Specify either channelId or channelName but not both

`--tabId [tabId]`
: The ID of the Microsoft Teams tab. Specify either tabId or tabName but not both

`--tabName [tabName]`
: The display name of the Microsoft Teams tab. Specify either tabId or tabName but not both

--8<-- "docs/cmd/_global.md"

## Remarks

You can only retrieve tabs for teams of which you are a member.

## Examples
  
Get a Microsoft Teams Tab with ID _1432c9da-8b9c-4602-9248-e0800f3e3f07_

```sh
m365 teams tab get --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --tabId 1432c9da-8b9c-4602-9248-e0800f3e3f07
```

Get a Microsoft Teams Tab with name _Tab Name_

```sh
m365 teams tab list --teamName "Team Name" --channelName "Channel Name" --tabName "Tab Name"
```
