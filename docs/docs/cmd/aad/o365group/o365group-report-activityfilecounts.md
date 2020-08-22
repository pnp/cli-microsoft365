# aad o365group report activityfilecounts

Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group

## Usage

```sh
aad o365group report activityfilecounts [options]
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

Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group for the last week

```sh
aad o365group report activityfilecounts --period D7
```

Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group for the last week and exports the report data in the specified path in text format

```sh
aad o365group report activityfilecounts --period D7 --output text > "o365groupactivityfilecounts.txt"
```

Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group for the last week and exports the report data in the specified path in json format

```sh
aad o365group report activityfilecounts --period D7 --output json > "o365groupactivityfilecounts.json"
```