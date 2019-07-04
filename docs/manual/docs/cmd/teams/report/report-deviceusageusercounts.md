# teams report deviceusageusercounts

Get the number of Microsoft Teams daily unique users by device type

## Usage

```sh
teams report deviceusageusercounts [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7|D30|D90|D180`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets the number of Microsoft Teams daily unique users by device type for the last week

```sh
teams report deviceusageusercounts --period D7
```