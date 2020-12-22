# outlook report mailappusageuserdetail

Gets details about which activities users performed on the various email apps

## Usage

```sh
m365 outlook report mailappusageuserdetail [options]
```

## Options

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is YYYY-MM-DD. Specify the date or period, but not both

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets details about which activities users performed on the various email apps for the last week

```sh
m365 outlook report mailappusageuserdetail --period D7
```

Gets details about which activities users performed on the various email apps for May 1st, 2019

```sh
m365 outlook report mailappusageuserdetail --date 2019-05-01
```

Gets details about which activities users performed on the various email apps for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailappusageuserdetail --period D7 --output text > "mailappusageuserdetail.txt"
```

Gets details about which activities users performed on the various email apps for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailappusageuserdetail --period D7 --output json > "mailappusageuserdetail.json"
```
