# spo report activityusercounts

Gets the trend in the number of active users

## Usage

```sh
m365 spo report activityusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

A user is considered active if he or she has executed a file activity (save, sync, modify, or share) or visited a page within the specified time period

## Examples

Gets the trend in the number of active users for the last week

```sh
m365 spo report activityusercounts --period D7
```

Gets the trend in the number of active users for the last week and exports the report data in the specified path in text format

```sh
m365 spo report activityusercounts --period D7 --output text > "activityusercounts.txt"
```

Gets the trend in the number of active users for the last week and exports the report data in the specified path in json format

```sh
m365 spo report activityusercounts --period D7 --output json > "activityusercounts.json"
```
