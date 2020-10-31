# spo report siteusagestorage

Gets the trend of storage allocated and consumed during the reporting period

## Usage

```sh
m365 spo report siteusagestorage [options]
```

## Options

`-h, --help`
: output usage information

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `text,json`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Gets the trend of storage allocated and consumed during the last week

```sh
m365 spo report siteusagestorage --period D7
```

Gets the trend of storage allocated and consumed during the last week and exports the report data in the specified path in text format

```sh
m365 spo report siteusagestorage --period D7 --output text > "siteusagestorage.txt"
```

Gets the trend of storage allocated and consumed during the last week and exports the report data in the specified path in json format

```sh
m365 spo report siteusagestorage --period D7 --output json > "siteusagestorage.json"
```
