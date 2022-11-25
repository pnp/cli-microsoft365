# onedrive report usageaccountcounts

Gets the trend in the number of active OneDrive for Business sites

## Usage

```sh
m365 onedrive report usageaccountcounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

Any site on which users viewed, modified, uploaded, downloaded, shared, or synced files is considered an active site

## Examples

Gets the trend in the number of active OneDrive for Business sites for the last week

```sh
m365 onedrive report usageaccountcounts --period D7
```

Gets the trend in the number of active OneDrive for Business sites for the last week and exports the report data in the specified path in text format

```sh
m365 onedrive report usageaccountcounts --period D7 --output text > "usageaccountcounts.txt"
```

Gets the trend in the number of active OneDrive for Business sites for the last week and exports the report data in the specified path in json format

```sh
m365 onedrive report usageaccountcounts --period D7 --output json > "usageaccountcounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Total": "69",
        "Active": "61",
        "Report Date": "2022-10-25",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Total": "69",
        "Active": "59",
        "Report Date": "2022-10-24",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Total": "69",
        "Active": "20",
        "Report Date": "2022-10-23",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Total": "69",
        "Active": "10",
        "Report Date": "2022-10-22",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Total": "69",
        "Active": "54",
        "Report Date": "2022-10-21",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Total": "69",
        "Active": "51",
        "Report Date": "2022-10-20",
        "Report Period": "7"
      },
      {
        "Report Refresh Date": "2022-10-25",
        "Site Type": "All",
        "Total": "69",
        "Active": "69",
        "Report Date": "2022-10-19",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```csv
    Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
    2022-10-25,All,69,61,2022-10-25,7
    2022-10-25,All,69,59,2022-10-24,7
    2022-10-25,All,69,20,2022-10-23,7
    2022-10-25,All,69,10,2022-10-22,7
    2022-10-25,All,69,54,2022-10-21,7
    2022-10-25,All,69,51,2022-10-20,7
    2022-10-25,All,69,69,2022-10-19,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
    2022-10-25,All,69,61,2022-10-25,7
    2022-10-25,All,69,59,2022-10-24,7
    2022-10-25,All,69,20,2022-10-23,7
    2022-10-25,All,69,10,2022-10-22,7
    2022-10-25,All,69,54,2022-10-21,7
    2022-10-25,All,69,51,2022-10-20,7
    2022-10-25,All,69,69,2022-10-19,7
    ```
