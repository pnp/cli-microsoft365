# spo report siteusagesitecounts

Gets the total number of files across all sites and the number of active files

## Usage

```sh
spo report siteusagesitecounts [options]
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

## Remarks

A file (user or system) is considered active if it has been saved, synced, modified, or shared within the specified time period.

## Examples

Gets the total number of files across all sites and the number of active files for the last week

```sh
spo report siteusagesitecounts --period D7
```

Gets the total number of files across all sites and the number of active files for the last week and exports the report data in the specified path in text format

```sh
spo report siteusagesitecounts --period D7 --output text > "siteusagesitecounts.txt"
```

Gets the total number of files across all sites and the number of active files for the last week and exports the report data in the specified path in json format

```sh
spo report siteusagesitecounts --period D7 --output json > "siteusagesitecounts.json"
```
