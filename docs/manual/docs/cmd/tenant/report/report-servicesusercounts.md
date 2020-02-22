# tenant report servicesusercount

Gets the count of users by activity type and service.

## Usage

```sh
tenant report servicesusercount [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7|D30|D90|D180`
`-f, --outputFile [outputFile]`|Path to the file where the Microsoft Teams daily unique users by device type report should be stored in
`-o, --output [output]`|Output type. `text|json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets the count of users by activity type and service for the last week

```sh
tenant report servicesusercount --period D7
```

Gets the count of users by activity type and service for the last week and exports the report data in the specified path in text format

```sh
tenant report servicesusercount --period D7 --output text --outputFile servicesusercount.txt
```

Gets the count of users by activity type and service for the last week and exports the report data in the specified path in json format

```sh
tenant report servicesusercount --period D7 --output json --outputFile servicesusercount.json
```
