# teams report useractivityuserdetail

Get details about Microsoft Teams user activity by user.

## Usage

```sh
teams report useractivityuserdetail [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7, D30, D90, D180`. Specify the `period` or `date`, but not both.
`-d, --date [date]`|The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`. Specify the `date` or `period`, but not both.
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets details about Microsoft Teams user activity by user for the last week

```sh
teams report useractivityuserdetail --period D7
```
Gets details about Microsoft Teams user activity by user for July 13, 2019

```sh
teams report useractivityuserdetail --date 2019-07-13
```