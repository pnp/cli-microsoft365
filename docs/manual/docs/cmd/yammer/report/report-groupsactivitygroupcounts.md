# yammer report groupsactivitygroupcounts

Gets the total number of groups that existed and how many included group conversation activity

## Usage

```sh
yammer report groupsactivitygroupcounts [options]
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

Gets the total number of groups that existed and how many included group conversation activity for the last week

```sh
yammer report groupsactivitygroupcounts --period D7
```

Gets the total number of groups that existed and how many included group conversation activity for the last week and exports the report data in the specified path in text format

```sh
yammer report groupsactivitygroupcounts --period D7 --output text > "groupsactivitygroupcounts.txt"
```

Gets the total number of groups that existed and how many included group conversation activity for the last week and exports the report data in the specified path in json format

```sh
yammer report groupsactivitygroupcounts --period D7 --output json > "groupsactivitygroupcounts.json"
```
