# graph teams funsettings list

Lists fun settings for the specified Microsoft Teams team

## Usage

```sh
graph teams funsettings list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team for which to list fun settings
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

List fun settings of a Microsoft Teams team

```sh
graph teams funsettings list --teamId 83cece1e-938d-44a1-8b86-918cf6151957
```