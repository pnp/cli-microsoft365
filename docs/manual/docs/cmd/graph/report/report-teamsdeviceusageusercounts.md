# graph report teamsdeviceusageusercounts

Get the number of Microsoft Teams daily unique users by device type

## Usage

```sh
graph report teamsdeviceusageusercounts [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7|D30|D90|D180`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To get the number of Microsoft Teams daily unique users by device type, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

## Examples

Gets the number of Microsoft Teams daily unique users by device type for the last week

```sh
graph report teamsdeviceusageusercounts --period D7
```