# teams report pstn calls

Get details about PSTN calls made within a given time period

## Usage

```sh
m365 teams report pstncalls [options]
```

## Options

`-h, --help`
: output usage information

`--fromDateTime <fromDateTime>`
: The start of time range to query. UTC, inclusive

`--toDateTime [toDateTime]`
: The end time range to query. UTC, inclusive. Defaults to current DateTime if omitted

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `text,json`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get details about PSTN calls made between 2020-10-31 and current DateTime

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31
```

Get details about PSTN calls made between 2020-10-31 and 2020-12-31 and exports the report data in the specified path in text format

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31 --toDateTime 2020-12-31 --output text > "pstncalls.txt"
```

Get details about PSTN calls made between 2020-10-31 and 2020-12-31 and exports the report data in the specified path in json format

```sh
m365 teams report pstncalls --fromDateTime 2020-10-31 --toDateTime 2020-12-31 --output json > "pstncalls.json"
```
