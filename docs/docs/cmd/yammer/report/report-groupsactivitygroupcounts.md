# yammer report groupsactivitygroupcounts

Gets the total number of groups that existed and how many included group conversation activity

## Usage

```sh
m365 yammer report groupsactivitygroupcounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the total number of groups that existed and how many included group conversation activity for the last week

```sh
m365 yammer report groupsactivitygroupcounts --period D7
```

Gets the total number of groups that existed and how many included group conversation activity for the last week and exports the report data in the specified path in text format

```sh
m365 yammer report groupsactivitygroupcounts --period D7 --output text > "groupsactivitygroupcounts.txt"
```

Gets the total number of groups that existed and how many included group conversation activity for the last week and exports the report data in the specified path in json format

```sh
m365 yammer report groupsactivitygroupcounts --period D7 --output json > "groupsactivitygroupcounts.json"
```
