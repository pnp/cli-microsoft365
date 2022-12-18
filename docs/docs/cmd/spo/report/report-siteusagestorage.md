# spo report siteusagestorage

Gets the trend of storage allocated and consumed during the reporting period

## Usage

```sh
m365 spo report siteusagestorage [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7`, `D30`, `D90`, `D180`.

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in.

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trend of storage allocated and consumed during the last week

```sh
m365 spo report siteusagestorage --period D7
```

Gets the trend of storage allocated and consumed during the last week and exports the report data in the specified path in text format

```sh
m365 spo report siteusagestorage --period D7 --output text > "siteusagestorage.txt"
```

Gets the trend of storage allocated and consumed during the last week and exports the report data in the specified path in json format

```sh
m365 spo report siteusagestorage --period D7 --output json > "siteusagestorage.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-26",
        "Site Type": "All",
        "Storage Used (Byte)": "2348104595",
        "Report Date": "2022-11-26",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Site Type,Storage Used (Byte),Report Date,Report Period
    2022-11-26,All,2348104595,2022-11-26,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Site Type,Storage Used (Byte),Report Date,Report Period
    2022-11-26,All,2348104595,2022-11-26,7
    ```
