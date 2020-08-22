# aad o365group report activitygroupcounts

Get the daily total number of groups and how many of them were active based on email conversations, Yammer posts, and SharePoint file activities.

## Usage

```sh
aad o365group report activitygroupcounts [options]
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

Get the daily total number of groups and how many of them were active based on activities for the last week

```sh
aad o365group report activitygroupcounts --period D7
```

Get the daily total number of groups and how many of them were active based on activities for the last week and exports the report data in the specified path in text format

```sh
aad o365group report activitygroupcounts --period D7 --output text > "o365groupactivitygroupcounts.txt"
```

Get the daily total number of groups and how many of them were active based on activities for the last week and exports the report data in the specified path in json format

```sh
aad o365group report activitygroupcounts --period D7 --output json > "o365groupactivitygroupcounts.json"
```