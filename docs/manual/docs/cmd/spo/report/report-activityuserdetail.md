# spo report activityuserdetail

Gets details about SharePoint activity by user.

## Usage

```sh
spo report activityuserdetail [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-d, --date [date]`|The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`. Specify the date or period, but not both.
`-p, --period [period]`|The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`
`-o, --output [output]`|Output type. `text,json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets details about SharePoint activity by user for the last week

```sh
spo report activityuserdetail --period D7
```

Gets details about SharePoint activity by user for May 1, 2019

```sh
spo report activityuserdetail --date 2019-05-01
```

Gets details about SharePoint activity by user for the last week and exports the report data in the specified path in text format

```sh
spo report activityuserdetail --period D7 --output text > "activityuserdetail.txt"
```

Gets details about SharePoint activity by user for the last week and exports the report data in the specified path in json format

```sh
spo report activityuserdetail --period D7 --output json > "activityuserdetail.json"
```
