# skype report activityusercounts

Gets the trends on how many unique users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions

## Usage

```sh
m365 skype report activityusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trends on how many unique users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week

```sh
m365 skype report activityusercounts --period D7
```

Gets the trends on how many unique users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week and exports the report data in the specified path in text format

```sh
m365 skype report activityusercounts --period D7 --output text > "activityusercounts.txt"
```

Gets the trends on how many unique users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week and exports the report data in the specified path in json format

```sh
m365 skype report activityusercounts --period D7 --output json > "activityusercounts.json"
```
