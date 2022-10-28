# onedrive report activityuserdetail

Gets details about OneDrive activity by user

## Usage

```sh
m365 onedrive report activityuserdetail [options]
```

## Options

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is YYYY-MM-DD. Specify the date or period, but not both`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets details about OneDrive activity by user for the last week

```sh
m365 onedrive report activityuserdetail --period D7
```

Gets details about OneDrive activity by user for May 1, 2019

```sh
m365 onedrive report activityuserdetail --date 2019-05-01
```

Gets details about OneDrive activity by user for the last week and exports the report data in the specified path in text format

```sh
m365 onedrive report activityuserdetail --period D7 --output text > "onedriveactivityuserdetail.txt"
```

Gets details about OneDrive activity by user for the last week and exports the report data in the specified path in json format

```sh
m365 onedrive report activityuserdetail --period D7 --output json > "onedriveactivityuserdetail.json"
```

## Response

=== "JSON"

```json
[
  {
    "Report Refresh Date": "2022-10-25",
    "User Principal Name": "C53803D3F8266856872FF413A333714F",
    "Is Deleted": "False",
    "Deleted Date": "",
    "Last Activity Date": "2022-10-25",
    "Viewed Or Edited File Count": "10",
    "Synced File Count": "125",
    "Shared Internally File Count": "6",
    "Shared Externally File Count": "0",
    "Assigned Products": "MICROSOFT 365 E3+POWER BI PREMIUM PER USER ADD-ON+POWER BI PRO+MICROSOFT POWER APPS PLAN 2 TRIAL+POWER BI (FREE)+AZURE ACTIVE DIRECTORY BASIC",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "User Principal Name": "C5E18087B685EA9F06A4517222C79046",
    "Is Deleted": "False",
    "Deleted Date": "",
    "Last Activity Date": "",
    "Viewed Or Edited File Count": "0",
    "Synced File Count": "0",
    "Shared Internally File Count": "0",
    "Shared Externally File Count": "0",
    "Assigned Products": "MICROSOFT 365 E3",
    "Report Period": "7"
  }
]
```

=== "Text"

    ``` text

Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Assigned Products,Report Period
2022-10-25,C53803D3F8266856872FF413A333714F,False,,2022-10-25,10,125,6,0,MICROSOFT 365 E3+POWER BI PREMIUM PER USER ADD-ON+POWER BI PRO+MICROSOFT POWER APPS PLAN 2 TRIAL+POWER BI (FREE)+AZURE ACTIVE DIRECTORY BASIC,7
2022-10-25,C5E18087B685EA9F06A4517222C79046,False,,,0,0,0,0,MICROSOFT 365 E3,7

````

=== "CSV"

    ``` text
Report Refresh Date,User Principal Name,Is Deleted,Deleted Date,Last Activity Date,Viewed Or Edited File Count,Synced File Count,Shared Internally File Count,Shared Externally File Count,Assigned Products,Report Period
2022-10-25,C53803D3F8266856872FF413A333714F,False,,2022-10-25,10,125,6,0,MICROSOFT 365 E3+POWER BI PREMIUM PER USER ADD-ON+POWER BI PRO+MICROSOFT POWER APPS PLAN 2 TRIAL+POWER BI (FREE)+AZURE ACTIVE DIRECTORY BASIC,7
2022-10-25,C5E18087B685EA9F06A4517222C79046,False,,,0,0,0,0,MICROSOFT 365 E3,7
````
