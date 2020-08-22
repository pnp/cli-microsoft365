# outlook report mailactivityuserdetail

Gets details about email activity users have performed

## Usage

```sh
outlook report mailactivityuserdetail [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-p, --period [period]`|The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`
`-d, --date [date]`|The date for which you would like to view the users who performed any activity. Supported date format is YYYY-MM-DD. Specify the date or period, but not both
`--query [query]`|JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples
`-o, --output [output]`|Output type. `text,json`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

## Examples

Gets details about email activity users have performed for the last week

```sh
outlook report mailactivityuserdetail --period D7
```

Gets details about email activity users have performed for May 1st, 2019

```sh
outlook report mailactivityuserdetail --date 2019-05-01
```

Gets details about email activity users have performed for the last week and exports the report data in the specified path in text format

```sh
outlook report mailactivityuserdetail --period D7 --output text > "mailactivityuserdetail.txt"
```

Gets details about email activity users have performed for the last week and exports the report data in the specified path in json format

```sh
outlook report mailactivityuserdetail --period D7 --output json > "mailactivityuserdetail.json"
```
