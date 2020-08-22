# aad o365group report activitycounts

Get the number of group activities across group workloads

## Usage

```sh
aad o365group report activitycounts [options]
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

Get the number of group activities across group workloads for the last week

```sh
aad o365group report activitycounts --period D7
```

Get the number of group activities across group workloads for the last week and exports the report data in the specified path in text format

```sh
aad o365group report activitycounts --period D7 --output text > "o365groupactivitycounts.txt"
```

Get the number of group activities across group workloads for the last week and exports the report data in the specified path in json format

```sh
aad o365group report activitycounts --period D7 --output json > "o365groupactivitycounts.json"
```