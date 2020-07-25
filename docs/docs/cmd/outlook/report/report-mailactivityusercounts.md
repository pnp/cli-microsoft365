# outlook report mailactivityusercounts

Enables you to understand trends on the number of unique users who are performing email activities like send, read, and receive

## Usage

```sh
m365 outlook report mailactivityusercounts [options]
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

Gets the trends on the number of unique users who are performing email activities like send, read, and receive for the last week

```sh
m365 outlook report mailactivityusercounts --period D7
```

Gets the trends on the number of unique users who are performing email activities like send, read, and receive for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailactivityusercounts --period D7 --output text > "mailactivityusercounts.txt"
```

Gets the trends on the number of unique users who are performing email activities like send, read, and receive for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailactivityusercounts --period D7 --output json > "mailactivityusercounts.json"
```
