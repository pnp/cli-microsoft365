# spo report activitypages

Gets the number of unique pages visited by users

## Usage

```sh
m365 spo report activitypages [options]
```

## Options

`-h, --help`
: output usage information

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

`-o, --output [output]`
: Output type. `text,json`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Gets the number of unique pages visited by users for the last week

```sh
m365 spo report activitypages --period D7
```

Gets the number of unique pages visited by users for the last week and exports the report data in the specified path in text format

```sh
m365 spo report activitypages --period D7 --output text > "activitypages.txt"
```

Gets the number of unique pages visited by users for the last week and exports the report data in the specified path in json format

```sh
m365 spo report activitypages --period D7 --output json > "activitypages.json"
```
