# teams report deviceusageuserdetail

Gets detail about Microsoft Teams device usage by user.

## Usage

```sh
teams report deviceusageuserdetail [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period [period]`|The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`
`-d, --date [date]`|The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`.
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `text,json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets information about Microsoft Teams device usage by user for the last week

```sh
teams report deviceusageuserdetail --period D7
```

Gets information about Microsoft Teams device usage by user for July 1, 2019

```sh
teams report deviceusageuserdetail --date 2019-07-01
```

Gets information about Microsoft Teams device usage by user for the last week and exports the report data in the specified path in text format

```sh
teams report deviceusageuserdetail --period D7 --output text > "deviceusageuserdetail.txt"
```

Gets information about Microsoft Teams device usage by user for the last week and exports the report data in the specified path in json format

```sh
teams report deviceusageuserdetail --period D7 --output json > "deviceusageuserdetail.json"
```
