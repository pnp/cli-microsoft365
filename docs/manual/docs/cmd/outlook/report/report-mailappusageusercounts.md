# outlook report mailappusageusercounts

Gets the count of unique users that connected to Exchange Online using any email app

## Usage

```sh
outlook report mailappusageusercounts [options]
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

Gets the count of unique users that connected to Exchange Online using any email app for the last week

```sh
outlook report mailappusageusercounts --period D7
```

Gets the count of unique users that connected to Exchange Online using any email app for the last week and exports the report data in the specified path in text format

```sh
outlook report mailappusageusercounts --period D7 --output text > "mailappusageusercounts.txt"
```

Gets the count of unique users that connected to Exchange Online using any email app for the last week and exports the report data in the specified path in json format

```sh
outlook report mailappusageusercounts --period D7 --output json > "mailappusageusercounts.json"
```
