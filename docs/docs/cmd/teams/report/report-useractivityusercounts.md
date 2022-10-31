# teams report useractivityusercounts

Get the number of Microsoft Teams users by activity type. The activity types are number of teams chat messages, private chat messages, calls, or meetings.

## Usage

```sh
m365 teams report useractivityusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft Teams users by activity type report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the number of Microsoft Teams users by activity type for the last week

```sh
m365 teams report useractivityusercounts --period D7
```

Gets the number of Microsoft Teams users by activity type for the last week and exports the report data in the specified path in text format

```sh
m365 teams report useractivityusercounts --period D7 --output text > "useractivityusercounts.txt"
```

Gets the number of Microsoft Teams users by activity type for the last week and exports the report data in the specified path in json format

```sh
m365 teams report useractivityusercounts --period D7 --output json > "useractivityusercounts.json"
```

## Response

=== "JSON"

    ``` json
    [
      {
        "Report Refresh Date": "2022-10-28",
        "Report Date": "2022-10-28",
        "Team Chat Messages": "0",
        "Private Chat Messages": "0",
        "Calls": "0",
        "Meetings": "0",
        "Other Actions": "0",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-28",
        "Report Date": "2022-10-27",
        "Team Chat Messages": "0",
        "Private Chat Messages": "0",
        "Calls": "0",
        "Meetings": "0",
        "Other Actions": "0",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-28",
        "Report Date": "2022-10-26",
        "Team Chat Messages": "1",
        "Private Chat Messages": "0",
        "Calls": "0",
        "Meetings": "0",
        "Other Actions": "0",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-28",
        "Report Date": "2022-10-25",
        "Team Chat Messages": "0",
        "Private Chat Messages": "0",
        "Calls": "0",
        "Meetings": "0",
        "Other Actions": "0",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-28",
        "Report Date": "2022-10-24",
        "Team Chat Messages": "0",
        "Private Chat Messages": "0",
        "Calls": "0",
        "Meetings": "0",
        "Other Actions": "0",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-28",
        "Report Date": "2022-10-23",
        "Team Chat Messages": "0",
        "Private Chat Messages": "0",
        "Calls": "0",
        "Meetings": "0",
        "Other Actions": "0",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-28",
        "Report Date": "2022-10-22",
        "Team Chat Messages": "0",
        "Private Chat Messages": "0",
        "Calls": "0",
        "Meetings": "0",
        "Other Actions": "0",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ``` text
    Report Refresh Date,Report Date,Team Chat Messages,Private Chat Messages,Calls,Meetings,Other Actions,Report Period
    2022-10-28,2022-10-28,0,0,0,0,0,7
    2022-10-28,2022-10-27,0,0,0,0,0,7
    2022-10-28,2022-10-26,1,0,0,0,0,7
    2022-10-28,2022-10-25,0,0,0,0,0,7
    2022-10-28,2022-10-24,0,0,0,0,0,7
    2022-10-28,2022-10-23,0,0,0,0,0,7
    2022-10-28,2022-10-22,0,0,0,0,0,
    ```

=== "CSV"

    ``` text
    Report Refresh Date,Report Date,Team Chat Messages,Private Chat Messages,Calls,Meetings,Other Actions,Report Period
    2022-10-28,2022-10-28,0,0,0,0,0,7
    2022-10-28,2022-10-27,0,0,0,0,0,7
    2022-10-28,2022-10-26,1,0,0,0,0,7
    2022-10-28,2022-10-25,0,0,0,0,0,7
    2022-10-28,2022-10-24,0,0,0,0,0,7
    2022-10-28,2022-10-23,0,0,0,0,0,7
    2022-10-28,2022-10-22,0,0,0,0,0,7
    ```
