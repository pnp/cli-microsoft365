# graph reports Microsoft Teams device usage by user

Gets detail about Microsoft Teams device usage by user

## Usage

```sh
graph reports teamsdeviceusageuserdetail [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period [period]`|Specify the length of time over which the report is aggregated. The supported values are `D7|D30|D90|D180`.
`-d, --date [date]`|Specify the date for which you would like to view the users who performed any activity. The supported date format is `YYYY-MM-DD`.
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to the Microsoft Graph, using the [graph login](../login.md) command.

## Remarks

To get details about Microsoft Teams device usage by user, you have to first log in to the Microsoft Graph using the [graph login](../login.md) command, eg. `graph login`.

Reports.Read.All permission is required to call this API.

## Examples

Gets detail about Microsoft Teams device usage by user for the length of time over which the report is aggregated

```sh
graph reports teamsdeviceusageuserdetail --period 'D7'
```
```sh
graph reports teamsdeviceusageuserdetail --period D7
```

Gets detail about Microsoft Teams device usage by user for date for which you would like to view the users who performed any activity

```sh
graph reports teamsdeviceusageuserdetail --date 2019-05-01
```
```sh
graph reports teamsdeviceusageuserdetail --date '2019-05-01'
```