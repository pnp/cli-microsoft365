# outlook report mailboxusagestorage

Gets the amount of mailbox storage used in your organization

## Usage

```sh
outlook report mailboxusagestorage [options]
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

Gets the amount of mailbox storage used in your organization for the last week

```sh
outlook report mailboxusagestorage --period D7
```

Gets the amount of mailbox storage used in your organization for the last week and exports the report data in the specified path in text format

```sh
outlook report mailboxusagestorage --period D7 --output text > "mailboxusagestorage.txt"
```

Gets the amount of mailbox storage used in your organization for the last week and exports the report data in the specified path in json format

```sh
outlook report mailboxusagestorage --period D7 --output json > "mailboxusagestorage.json"
```
