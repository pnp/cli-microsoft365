# graph teams membersettings list

Lists member settings for a Microsoft Teams team

## Usage

```sh
graph teams membersettings list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId`|The ID of the team for which to get the member settings
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Get member settings for a Microsoft Teams team

```sh
graph teams membersettings list --teamId 2609af39-7775-4f94-a3dc-0dd67657e900
```