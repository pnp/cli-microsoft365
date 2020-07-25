# teams report useractivityusercounts

Get the number of Microsoft Teams users by activity type. The activity types are number of teams chat messages, private chat messages, calls, or meetings.

## Usage

```sh
m365 teams report useractivityusercounts [options]
```

## Options

`-h, --help`
: output usage information

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft Teams users by activity type report should be stored in

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `text,json`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

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
