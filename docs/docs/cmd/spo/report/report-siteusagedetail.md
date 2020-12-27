# spo report siteusagedetail

Gets details about SharePoint site usage

## Usage

```sh
m365 spo report siteusagedetail [options]
```

## Options

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`. Specify the date or period, but not both.

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets details about SharePoint site usage for the last week

```sh
m365 spo report siteusagedetail --period D7
```

Gets details about SharePoint site usage for May 1, 2019

```sh
m365 spo report siteusagedetail --date 2019-05-01
```

Gets details about SharePoint site usage for the last week and exports the report data in the specified path in text format

```sh
m365 spo report siteusagedetail --period D7 --output text > "siteusagedetail.txt"
```

Gets details about SharePoint site usage for the last week and exports the report data in the specified path in json format

```sh
m365 spo report siteusagedetail --period D7 --output json > "siteusagedetail.json"
```
