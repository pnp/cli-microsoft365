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

## Response

=== "JSON"

    ``` json
    [
      {
        "Report Refresh Date": "2022-10-28",
        "User Id": "00000000-0000-0000-0000-000000000000",
        "User Principal Name": "6E55A185B405B6F2A6804BB7897C8AAB",
        "Last Activity Date": "2022-10-26",
        "Is Deleted": "False",
        "Deleted Date": "",
        "Assigned Products": "MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)+MICROSOFT POWER AUTOMATE FREE",
        "Team Chat Message Count": "5",
        "Private Chat Message Count": "0",
        "Call Count": "0",
        "Meeting Count": "0",
        "Meetings Organized Count": "0",
        "Meetings Attended Count": "0",
        "Ad Hoc Meetings Organized Count": "0",
        "Ad Hoc Meetings Attended Count": "0",
        "Scheduled One-time Meetings Organized Count": "0",
        "Scheduled One-time Meetings Attended Count": "0",
        "Scheduled Recurring Meetings Organized Count": "0",
        "Scheduled Recurring Meetings Attended Count": "0",
        "Audio Duration": "PT0S",
        "Video Duration": "PT0S",
        "Screen Share Duration": "PT0S",
        "Audio Duration In Seconds": "0",
        "Video Duration In Seconds": "0",
        "Screen Share Duration In Seconds": "0",
        "Has Other Action": "No",
        "Urgent Messages": "0",
        "Post Messages": "2",
        "Tenant Display Name": "CONTOSO",
        "Shared Channel Tenant Display Names": "",
        "Reply Messages": "3",
        "Is Licensed": "Yes",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ``` text
    Report Refresh Date,User Id,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Meetings Organized Count,Meetings Attended Count,Ad Hoc Meetings Organized Count,Ad Hoc Meetings Attended Count,Scheduled One-time Meetings Organized Count,Scheduled One-time Meetings Attended Count,Scheduled Recurring Meetings Organized Count,Scheduled Recurring Meetings Attended Count,Audio Duration,Video Duration,Screen Share Duration,Audio Duration In Seconds,Video Duration In Seconds,Screen Share Duration In Seconds,Has Other Action,Urgent Messages,Post Messages,Tenant Display Name,Shared Channel Tenant Display Names,Reply Messages,Is Licensed,Report Period
    2022-10-28,00000000-0000-0000-0000-000000000000,6E55A185B405B6F2A6804BB7897C8AAB,2022-10-26,False,,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)+MICROSOFT POWER AUTOMATE FREE,5,0,0,0,0,0,0,0,0,0,0,0,PT0S,PT0S,PT0S,0,0,0,No,0,2,CONTOSO,,3,Yes,7
    ```

=== "CSV"

    ``` text
    Report Refresh Date,User Id,User Principal Name,Last Activity Date,Is Deleted,Deleted Date,Assigned Products,Team Chat Message Count,Private Chat Message Count,Call Count,Meeting Count,Meetings Organized Count,Meetings Attended Count,Ad Hoc Meetings Organized Count,Ad Hoc Meetings Attended Count,Scheduled One-time Meetings Organized Count,Scheduled One-time Meetings Attended Count,Scheduled Recurring Meetings Organized Count,Scheduled Recurring Meetings Attended Count,Audio Duration,Video Duration,Screen Share Duration,Audio Duration In Seconds,Video Duration In Seconds,Screen Share Duration In Seconds,Has Other Action,Urgent Messages,Post Messages,Tenant Display Name,Shared Channel Tenant Display Names,Reply Messages,Is Licensed,Report Period
    2022-10-28,00000000-0000-0000-0000-000000000000,6E55A185B405B6F2A6804BB7897C8AAB,2022-10-26,False,,MICROSOFT 365 E5 DEVELOPER (WITHOUT WINDOWS AND AUDIO CONFERENCING)+MICROSOFT POWER AUTOMATE FREE,5,0,0,0,0,0,0,0,0,0,0,0,PT0S,PT0S,PT0S,0,0,0,No,0,2,CONTOSO,,3,Yes,7
    ```
