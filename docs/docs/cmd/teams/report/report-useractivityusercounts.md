# teams report useractivityusercounts

Get the number of Microsoft Teams users by activity type. The activity types are number of teams chat messages, private chat messages, calls, or meetings.

## Usage

```sh
m365 teams report useractivityusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft Teams users by activity type report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the number of Microsoft Teams users by activity type for the last week

```sh
m365 teams report useractivityusercounts --period D7
```

Gets the number of Microsoft Teams users by activity type for the last week and exports the report data in the specified path in text format

```sh
m365 teams report useractivityusercounts --period D7 --output text > "useractivityusercounts.txt"
```

Gets the number of Microsoft Teams users by activity type for the last week and exports the report data in the specified path in json format

```sh
m365 teams report useractivityusercounts --period D7 --output json > "useractivityusercounts.json"
```
