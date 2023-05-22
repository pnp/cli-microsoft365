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
        "Report Refresh Date": "2023-05-20",
        "Site Type": "All",
        "Total": "1",
        "Active": "0",
        "Report Date": "2023-05-20",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
    2023-05-20,All,1,0,2023-05-20,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
    2023-05-20,All,1,0,2023-05-20,7
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
    2023-05-20,All,1,0,2023-05-20,7
    ```
