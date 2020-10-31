# aad o365group report activitycounts

Get the number of group activities across group workloads

## Usage

```sh
m365 aad o365group report activitycounts [options]
```

## Options

`-h, --help`
: output usage information

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft 365 Groups activities across group workloads report should be stored in

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `text,json`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Examples

Get the number of group activities across group workloads for the last week

```sh
m365 aad o365group report activitycounts --period D7
```

Get the number of group activities across group workloads for the last week and exports the report data in the specified path in text format

```sh
m365 aad o365group report activitycounts --period D7 --output text > "o365groupactivitycounts.txt"
```

Get the number of group activities across group workloads for the last week and exports the report data in the specified path in json format

```sh
m365 aad o365group report activitycounts --period D7 --output json > "o365groupactivitycounts.json"
```