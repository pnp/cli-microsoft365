# aad o365group report activityfilecounts

Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group

## Usage

```sh
m365 aad o365group report activityfilecounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft 365 Groups activities across group workloads report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group for the last week

```sh
m365 aad o365group report activityfilecounts --period D7
```

Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group for the last week and exports the report data in the specified path in text format

```sh
m365 aad o365group report activityfilecounts --period D7 --output text > "o365groupactivityfilecounts.txt"
```

Get the total number of files and how many of them were active across all group sites associated with an Microsoft 365 Group for the last week and exports the report data in the specified path in json format

```sh
m365 aad o365group report activityfilecounts --period D7 --output json > "o365groupactivityfilecounts.json"
```
