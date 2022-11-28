# spo report siteusagedetail

Gets details about SharePoint site usage

## Usage

```sh
m365 spo report siteusagedetail [options]
```

## Options

`-d, --date [date]`
: The date for which you would like to view the users who performed any activity. Supported date format is `YYYY-MM-DD`. Specify either `date` or `period`, but not both.

`-p, --period [period]`
: The length of time over which the report is aggregated. Supported values `D7`, `D30`, `D90`, `D180`. Specify either `date` or `period`, but not both.

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in.

--8<-- "docs/cmd/_global.md"

## Remarks

As this report is only available for the past 28 days, date parameter value should be a date from that range.

## Examples

Gets details about SharePoint site usage for the last week

```sh
m365 spo report siteusagedetail --period D7
```

Gets details about SharePoint site usage for May 1, 2019

```sh
m365 spo report siteusagedetail --date 2019-05-01
```

Gets details about SharePoint site usage for the last week and exports the report data in the specified path in text format

```sh
m365 spo report siteusagedetail --period D7 --output text > "siteusagedetail.txt"
```

Gets details about SharePoint site usage for the last week and exports the report data in the specified path in json format

```sh
m365 spo report siteusagedetail --period D7 --output json > "siteusagedetail.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-26",
        "Site Id": "002caf24-2672-4bc4-bc41-2811d5dba62c",
        "Site URL": "https://contoso.sharepoint.com/sites/ElMontePythons",
        "Owner Display Name": "El Monte Pythons Owners",
        "Is Deleted": "False",
        "Last Activity Date": "",
        "File Count": "2",
        "Active File Count": "0",
        "Page View Count": "0",
        "Visited Page Count": "0",
        "Storage Used (Byte)": "1475974",
        "Storage Allocated (Byte)": "27487790694400",
        "Root Web Template": "Group",
        "Owner Principal Name": "ElMontePythons@contoso.onmicrosoft.com",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Site Id,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Page View Count,Visited Page Count,Storage Used (Byte),Storage Allocated (Byte),Root Web Template,Owner Principal Name,Report Period
    2022-11-26,002caf24-2672-4bc4-bc41-2811d5dba62c,https://contoso.sharepoint.com/sites/ElMontePythons,El Monte Pythons Owners,False,,2,0,0,0,1475974,27487790694400,Group,ElMontePythons@contoso.onmicrosoft.com,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Site Id,Site URL,Owner Display Name,Is Deleted,Last Activity Date,File Count,Active File Count,Page View Count,Visited Page Count,Storage Used (Byte),Storage Allocated (Byte),Root Web Template,Owner Principal Name,Report Period
    2022-11-26,002caf24-2672-4bc4-bc41-2811d5dba62c,https://contoso.sharepoint.com/sites/ElMontePythons,El Monte Pythons Owners,False,,2,0,0,0,1475974,27487790694400,Group,ElMontePythons@contoso.onmicrosoft.com,7
    ```
