# graph teams tab list

Lists tabs in the specified Microsoft Teams channel

## Usage

```sh
graph teams tab list [options]
```

## Options

Option|Description
------|-----------
`--help`| output usage information
`-i, --teamId <teamId>`|The ID of the team of the specific channel
`-c, --channelId <channelId>`|The ID of the channel for which to list tabs
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command

## Remarks

To list the tabs in a Microsoft Teams channel, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples
  
List the channels in a specified Microsoft Teams team

```sh
graph teams channel list --teamId 00000000-0000-0000-0000-000000000000
```

## Remarks:

To list available tabs in a specific Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md)command, eg. `graph login`.

You can only see the tab list of a team you are a member of.

The tabs Conversations and Files are present in every team and therefor not provided in the response from the graph call. The command uses Microsoft Graph to retrive the tab information. More details on the underlying graph endpoint can be found at <https://docs.microsoft.com/en-us/graph/api/teamstab-list?view=graph-rest-1.0>

## Examples:
  
List all tabs in a Microsoft Teams channel

```sh
graph teams tab list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype
```
    
Include all the values from the tab configuration and associated teams app

```sh
graph teams tab list --teamId 00000000-0000-0000-0000-000000000000 --channelId 19:00000000000000000000000000000000@thread.skype --output json
```