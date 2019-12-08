# spo report onedriveusageaccountdetail

Gets details about OneDrive usage by account

## Usage

```sh
spo report onedriveusageaccountdetail [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period [period]`|The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`
`-d, --date [date]`|The date for which you would like to view the users who performed any activity. Supported date format is YYYY-MM-DD. Specify the date or period, but not both`
`-f, --outputFile [outputFile]`|Path to the file where the report should be stored in
`-o, --output [output]`|Output type. `text,json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets details about OneDrive usage by account for the last week

```sh
spo report onedriveusageaccountdetail --period D7
```

Gets details about OneDrive usage by account for May 1, 2019

```sh
spo report onedriveusageaccountdetail --date 2019-05-01
```

Gets details about OneDrive usage by account for the last week and exports the report data in the specified path in text format

```sh
spo report onedriveusageaccountdetail --period D7 --output text --outputFile 'C:/report.txt'
```

Gets details about OneDrive usage by account for the last week and exports the report data in the specified path in json format

```sh
spo report onedriveusageaccountdetail --period D7 --output json --outputFile 'C:/report.json'
```
