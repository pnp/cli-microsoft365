# outlook report mailactivitycounts

Enables you to understand the trends of email activity (like how many were sent, read, and received) in your organization

## Usage

```sh
m365 outlook report mailactivitycounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week

```sh
m365 outlook report mailactivitycounts --period D7
```

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailactivitycounts --period D7 --output text > "mailactivitycounts.txt"
```

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailactivitycounts --period D7 --output json > "mailactivitycounts.json"
```
