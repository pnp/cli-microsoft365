# aad o365group report activitystorage

Get the total storage used across all group mailboxes and group sites

## Usage

```sh
aad o365group report activitystorage  [options]
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

Get the total storage used across all group mailboxes and group sites for the last week

```sh
aad o365group report activitystorage --period D7
```

Get the total storage used across all group mailboxes and group sites for the last week and exports the report data in the specified path in text format

```sh
aad o365group report activitystorage --period D7 --output text > "o365groupactivitystorage.txt"
```

Get the total storage used across all group mailboxes and group sites for the last week and exports the report data in the specified path in json format

```sh
aad o365group report activitystorage --period D7 --output json > "o365groupactivitystorage.json"
```
