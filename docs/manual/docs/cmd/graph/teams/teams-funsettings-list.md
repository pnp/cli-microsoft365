# graph teams funsettings list

Lists all funsettings from a Microsoft Teams team

## Usage

```sh
graph teams funsettings list [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --teamId <teamId>`|The ID of the team
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

### Remarks

To get the funsettings of a Microsoft Teams team, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.


## Examples

Retrieve the funsettings of a Microsoft Teams team

```sh
graph teams funsettings list --teamId 83cece1e-938d-44a1-8b86-918cf6151957
```