# aad o365group report activitygroupcounts

Get the daily total number of groups and how many of them were active based on email conversations, Yammer posts, and SharePoint file activities.

## Usage

```sh
m365 aad o365group report activitygroupcounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft 365 Groups activities report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Get the daily total number of groups and how many of them were active based on activities for the last week

```sh
m365 aad o365group report activitygroupcounts --period D7
```

Get the daily total number of groups and how many of them were active based on activities for the last week and exports the report data in the specified path in text format

```sh
m365 aad o365group report activitygroupcounts --period D7 --output text > "o365groupactivitygroupcounts.txt"
```

Get the daily total number of groups and how many of them were active based on activities for the last week and exports the report data in the specified path in json format

```sh
m365 aad o365group report activitygroupcounts --period D7 --output json > "o365groupactivitygroupcounts.json"
```
