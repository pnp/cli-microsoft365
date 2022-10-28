# onedrive report activityusercounts

Gets the trend in the number of active OneDrive users

## Usage

```sh
m365 onedrive report activityusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trend in the number of active OneDrive users for the last week

```sh
m365 onedrive report activityusercounts --period D7
```

Gets the trend in the number of active OneDrive users for the last week and exports the report data in the specified path in text format

```sh
m365 onedrive report activityusercounts --period D7 --output text > "activityusercounts.txt"
```

Gets the trend in the number of active OneDrive users for the last week and exports the report data in the specified path in json format

```sh
m365 onedrive report activityusercounts --period D7 --output json > "activityusercounts.json"
```

## Response

=== "JSON"

```json
[
  {
    "Report Refresh Date": "2022-10-25",
    "Viewed Or Edited": "20",
    "Synced": "8",
    "Shared Internally": "5",
    "Shared Externally": "",
    "Report Date": "2022-10-25",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "Viewed Or Edited": "24",
    "Synced": "15",
    "Shared Internally": "5",
    "Shared Externally": "",
    "Report Date": "2022-10-24",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "Viewed Or Edited": "4",
    "Synced": "3",
    "Shared Internally": "",
    "Shared Externally": "",
    "Report Date": "2022-10-23",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "Viewed Or Edited": "4",
    "Synced": "1",
    "Shared Internally": "",
    "Shared Externally": "",
    "Report Date": "2022-10-22",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "Viewed Or Edited": "11",
    "Synced": "13",
    "Shared Internally": "1",
    "Shared Externally": "",
    "Report Date": "2022-10-21",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "Viewed Or Edited": "22",
    "Synced": "12",
    "Shared Internally": "8",
    "Shared Externally": "1",
    "Report Date": "2022-10-20",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "Viewed Or Edited": "20",
    "Synced": "12",
    "Shared Internally": "4",
    "Shared Externally": "",
    "Report Date": "2022-10-19",
    "Report Period": "7"
  }
]
```

=== "Text"

    ``` text

Report Refresh Date,Viewed Or Edited,Synced,Shared Internally,Shared Externally,Report Date,Report Period
2022-10-25,20,8,5,,2022-10-25,7
2022-10-25,24,15,5,,2022-10-24,7
2022-10-25,4,3,,,2022-10-23,7
2022-10-25,4,1,,,2022-10-22,7
2022-10-25,11,13,1,,2022-10-21,7
2022-10-25,22,12,8,1,2022-10-20,7
2022-10-25,20,12,4,,2022-10-19,7

````

=== "CSV"

    ``` text
Report Refresh Date,Viewed Or Edited,Synced,Shared Internally,Shared Externally,Report Date,Report Period
2022-10-25,20,8,5,,2022-10-25,7
2022-10-25,24,15,5,,2022-10-24,7
2022-10-25,4,3,,,2022-10-23,7
2022-10-25,4,1,,,2022-10-22,7
2022-10-25,11,13,1,,2022-10-21,7
2022-10-25,22,12,8,1,2022-10-20,7
2022-10-25,20,12,4,,2022-10-19,7
````
