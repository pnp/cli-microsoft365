# tenant report activeusercounts

Gets the count of daily active users in the reporting period by product.

## Usage

```sh
m365 tenant report activeusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the Microsoft Teams daily unique users by device type report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the count of daily active users in the reporting period by product for the last week

```sh
m365 tenant report activeusercounts --period D7
```

Gets the count of daily active users in the reporting period by product for the last week and exports the report data in the specified path in text format

```sh
m365 tenant report activeusercounts --period D7 --output text > "activeusercounts.txt"
```

Gets the count of daily active users in the reporting period by product for the last week and exports the report data in the specified path in json format

```sh
m365 tenant report activeusercounts --period D7 --output json > "activeusercounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-10-25",
        "Office 365": "1",
        "Exchange": "5",
        "OneDrive": "4",
        "SharePoint": "3",
        "Skype For Business": "2",
        "Yammer": "3",
        "Teams": "1",
        "Report Date": "2022-10-19",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Office 365,Exchange,OneDrive,SharePoint,Skype For Business,Yammer,Teams,Report Date,Report Period
    2022-10-25,1,5,4,3,2,3,1,2022-10-19,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Office 365,Exchange,OneDrive,SharePoint,Skype For Business,Yammer,Teams,Report Date,Report Period
    2022-10-25,1,5,4,3,2,3,1,2022-10-19,7
    ```
