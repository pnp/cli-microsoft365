# yammer report deviceusageuserdetail

Gets details about Yammer device usage by user

## Usage

```sh
m365 yammer report deviceusageuserdetail [options]
```

## Options

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`.

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets details about Yammer device usage by user for the last week

```sh
m365 yammer report deviceusageuserdetail --period D7
```

Gets details about Yammer device usage by user for July 1, 2019

```sh
m365 yammer report deviceusageuserdetail --date 2019-07-01
```

Gets details about Yammer device usage by user for the last week and exports the report data in the specified path in text format

```sh
m365 yammer report deviceusageuserdetail --period D7 --output text > "deviceusageuserdetail.txt"
```

Gets details about Yammer device usage by user for the last week and exports the report data in the specified path in json format

```sh
m365 yammer report deviceusageuserdetail --period D7 --output json > "deviceusageuserdetail.json"
```
