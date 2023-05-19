# spo report activityuserdetail

Gets details about SharePoint activity by user.

## Usage

```sh
m365 spo report activityuserdetail [options]
```

## Options

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`. Specify either `date` or `period`, but not both.

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7`, `D30`, `D90`, `D180`. Specify either `date` or `period`, but not both.

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft Teams device usage by user report should be stored in.

--8<-- "docs/cmd/_global.md"

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets details about SharePoint activity by user for the last week

```sh
m365 spo report activityuserdetail --period D7
```

Gets details about SharePoint activity by user for May 1, 2019

```sh
m365 spo report activityuserdetail --date 2019-05-01
```

Gets details about SharePoint activity by user for the last week and exports the report data in the specified path in text format

```sh
m365 spo report activityuserdetail --period D7 --output text > "activityuserdetail.txt"
```

Gets details about SharePoint activity by user for the last week and exports the report data in the specified path in json format

```sh
m365 spo report activityuserdetail --period D7 --output json > "activityuserdetail.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-26",
        "User Principal Name": "Amanda.Powell@contoso.onmicrosoft.com",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Last Activity Date": "2022-09-08",
        "Viewed Or Edited File Count": "0",
        "Synced File Count": "0",
        "Shared Internally File Count": "0",
        "Shared Externally File Count": "0",
        "Visited Page Count": "0",
        "Assigned Products": "MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)+MICROSOFT POWER AUTOMATE FREE",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Visited Page Count,Assigned Products,Report Period
    2022-11-26,Amanda.Powell@contoso.onmicrosoft.com,False,,2022-09-08,0,0,0,0,0,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)+MICROSOFT POWER AUTOMATE FREE,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Visited Page Count,Assigned Products,Report Period
    2022-11-26,Amanda.Powell@contoso.onmicrosoft.com,False,,2022-09-08,0,0,0,0,0,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)+MICROSOFT POWER AUTOMATE FREE,7
    ```

=== "Markdown"

    ```md
    Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Visited Page Count,Assigned Products,Report Period
    2023-05-04,6B56E500AC8309BD90D212680A2B9C03,False,,2023-05-04,16,0,0,0,36,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)+MICROSOFT POWER AUTOMATE FREE,7
    ```
