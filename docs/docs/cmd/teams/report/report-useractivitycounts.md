# teams report useractivitycounts

Get the number of Microsoft Teams activities by activity type. The activity types are team chat messages, private chat messages, calls, and meetings.

## Usage

```sh
teams report useractivitycounts [options]
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

Gets the number of Microsoft Teams activities by activity type for the last week

```sh
teams report useractivitycounts --period D7
```

Gets the number of Microsoft Teams activities by activity type for the last week and exports the report data in the specified path in text format

```sh
teams report useractivitycounts --period D7 --output text > "useractivitycounts.txt"
```

Gets the number of Microsoft Teams activities by activity type for the last week and exports the report data in the specified path in json format

```sh
teams report useractivitycounts --period D7 --output json > "useractivitycounts.json"
```
