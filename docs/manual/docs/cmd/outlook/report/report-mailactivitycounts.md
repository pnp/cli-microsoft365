# outlook report mailactivitycounts

Enables you to understand the trends of email activity (like how many were sent, read, and received) in your organization

## Usage

```sh
outlook report mailactivitycounts [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period <period>`|The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `text,json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week

```sh
outlook report mailactivitycounts --period D7
```

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week and exports the report data in the specified path in text format

```sh
outlook report mailactivitycounts --period D7 --output text > "mailactivitycounts.txt"
```

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week and exports the report data in the specified path in json format

```sh
outlook report mailactivitycounts --period D7 --output json > "mailactivitycounts.json"
```
