# skype report activitycounts

Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions

## Usage

```sh
skype report activitycounts [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `text,json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week

```sh
skype report activitycounts --period D7
```

Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week and exports the report data in the specified path in text format

```sh
skype report activitycounts --period D7 --output text > "activitycounts.txt"
```

Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week and exports the report data in the specified path in json format

```sh
skype report activitycounts --period D7 --output json > "activitycounts.json"
```
