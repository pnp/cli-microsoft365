# tenant report activeusercounts

Gets the count of daily active users in the reporting period by product.

## Usage

```sh
tenant report activeusercounts [options]
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

Gets the count of daily active users in the reporting period by product for the last week

```sh
tenant report activeusercounts --period D7
```

Gets the count of daily active users in the reporting period by product for the last week and exports the report data in the specified path in text format

```sh
tenant report activeusercounts --period D7 --output text > "activeusercounts.txt"
```

Gets the count of daily active users in the reporting period by product for the last week and exports the report data in the specified path in json format

```sh
tenant report activeusercounts --period D7 --output json > "activeusercounts.json"
```
