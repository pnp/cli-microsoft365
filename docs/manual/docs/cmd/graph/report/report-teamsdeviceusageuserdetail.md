# graph report teamsdeviceusageuserdetail

Gets detail about Microsoft Teams device usage by user

## Usage

```sh
graph report teamsdeviceusageuserdetail [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period [period]`|The length of time over which the report is aggregated. Supported values `D7|D30|D90|D180`
`-d, --date [date]`|The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`.
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To get details about Microsoft Teams device usage by user, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets information about Microsoft Teams device usage by user for the last week

```sh
graph report teamsdeviceusageuserdetail --period D7
```

Gets information about Microsoft Teams device usage by user for May 1, 2019

```sh
graph report teamsdeviceusageuserdetail --date 2019-05-28
```
