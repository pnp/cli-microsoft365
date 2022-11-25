# onedrive report activityfilecounts

Gets the number of unique, licensed users that performed file interactions against any OneDrive account

## Usage

```sh
m365 onedrive report activityfilecounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the number of unique, licensed users that performed file interactions against any OneDrive account for the last week

```sh
m365 onedrive report activityfilecounts --period D7
```

Gets the number of unique, licensed users that performed file interactions against any OneDrive account for the last week and exports the report data in the specified path in text format

```sh
m365 onedrive report activityfilecounts --period D7 --output text > "activityfilecounts.txt"
```

Gets the number of unique, licensed users that performed file interactions against any OneDrive account for the last week and exports the report data in the specified path in json format

```sh
m365 onedrive report activityfilecounts --period D7 --output json > "activityfilecounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-10-25",
        "Viewed Or Edited": "69",
        "Synced": "22",
        "Shared Internally": "7",
        "Shared Externally": "",
        "Report Date": "2022-10-25",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Viewed Or Edited": "61",
        "Synced": "352",
        "Shared Internally": "7",
        "Shared Externally": "",
        "Report Date": "2022-10-24",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Viewed Or Edited": "87",
        "Synced": "91",
        "Shared Internally": "",
        "Shared Externally": "",
        "Report Date": "2022-10-23",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Viewed Or Edited": "9",
        "Synced": "10",
        "Shared Internally": "",
        "Shared Externally": "",
        "Report Date": "2022-10-22",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Viewed Or Edited": "38",
        "Synced": "61",
        "Shared Internally": "4",
        "Shared Externally": "",
        "Report Date": "2022-10-21",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Viewed Or Edited": "40",
        "Synced": "23",
        "Shared Internally": "9",
        "Shared Externally": "2",
        "Report Date": "2022-10-20",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Viewed Or Edited": "201",
        "Synced": "105",
        "Shared Internally": "6",
        "Shared Externally": "",
        "Report Date": "2022-10-19",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```csv
    Report Refresh Date,Viewed Or Edited,Synced,Shared Internally,Shared Externally,Report Date,Report Period
    2022-10-25,69,22,7,,2022-10-25,7
    2022-10-25,61,352,7,,2022-10-24,7
    2022-10-25,87,91,,,2022-10-23,7
    2022-10-25,9,10,,,2022-10-22,7
    2022-10-25,38,61,4,,2022-10-21,7
    2022-10-25,40,23,9,2,2022-10-20,7
    2022-10-25,201,105,6,,2022-10-19,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Viewed Or Edited,Synced,Shared Internally,Shared Externally,Report Date,Report Period
    2022-10-25,69,22,7,,2022-10-25,7
    2022-10-25,61,352,7,,2022-10-24,7
    2022-10-25,87,91,,,2022-10-23,7
    2022-10-25,9,10,,,2022-10-22,7
    2022-10-25,38,61,4,,2022-10-21,7
    2022-10-25,40,23,9,2,2022-10-20,7
    2022-10-25,201,105,6,,2022-10-19,7
    ```
