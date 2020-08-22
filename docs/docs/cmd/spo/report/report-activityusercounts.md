# spo report activityusercounts

Gets the trend in the number of active users

## Usage

```sh
spo report activityusercounts [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`
`-o, --output [output]`|Output type. `text,json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

A user is considered active if he or she has executed a file activity (save, sync, modify, or share) or visited a page within the specified time period

## Examples

Gets the trend in the number of active users for the last week

```sh
spo report activityusercounts --period D7
```

Gets the trend in the number of active users for the last week and exports the report data in the specified path in text format

```sh
spo report activityusercounts --period D7 --output text > "activityusercounts.txt"
```

Gets the trend in the number of active users for the last week and exports the report data in the specified path in json format

```sh
spo report activityusercounts --period D7 --output json > "activityusercounts.json"
```
