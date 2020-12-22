# teams report useractivityuserdetail

Get details about Microsoft Teams user activity by user.

## Usage

```sh
m365 teams report useractivityuserdetail [options]
```

## Options

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`. Specify the `period` or `date`, but not both.

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`. Specify the `date` or `period`, but not both.

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft Teams user activity by user report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets details about Microsoft Teams user activity by user for the last week

```sh
m365 teams report useractivityuserdetail --period D7
```

Gets details about Microsoft Teams user activity by user for July 13, 2019

```sh
m365 teams report useractivityuserdetail --date 2019-07-13
```

Gets details about Microsoft Teams user activity by user for the last week and exports the report data in the specified path in text format

```sh
m365 teams report useractivityuserdetail --period D7 --output text > "useractivityuserdetail.txt"
```

Gets details about Microsoft Teams user activity by user for the last week and exports the report data in the specified path in json format

```sh
m365 teams report useractivityuserdetail --period D7 --output json > "useractivityuserdetail.json"
```
