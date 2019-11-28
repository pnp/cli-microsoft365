# outlook report getmailboxusagestorage

Get the amount of storage used in your organization. 

## Usage

```sh
outlook report getmailboxusagestorage [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7|D30|D90|D180`
`-f, --outputFile [outputFile]`|Path to the file where the report should be stored in
`-o, --output [output]`|Output type. `text|json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets the amount of storage used in your organization for the last week

```sh
outlook report getmailboxusagestorage --period D7
```

Gets the amount of storage used in your organization for the last week and exports the report data in the specified path in text format

```sh
outlook report getmailboxusagestorage --period D7 --output text --outputFile 'C:/report.txt'
```

Gets the amount of storage used in your organization for the last week and exports the report data in the specified path in json format

```sh
outlook report getmailboxusagestorage --period D7 --output json --outputFile 'C:/report.json'
```
