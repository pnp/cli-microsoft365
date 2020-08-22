# onerive report activityusercounts

Gets the trend in the number of active OneDrive users

## Usage

```sh
onedrive report activityusercounts [options]
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

Gets the trend in the number of active OneDrive users for the last week

```sh
onedrive report activityusercounts --period D7
```

Gets the trend in the number of active OneDrive users for the last week and exports the report data in the specified path in text format

```sh
onedrive report activityusercounts --period D7 --output text > "activityusercounts.txt"
```

Gets the trend in the number of active OneDrive users for the last week and exports the report data in the specified path in json format

```sh
onedrive report activityusercounts --period D7 --output json > "activityusercounts.json"
```
