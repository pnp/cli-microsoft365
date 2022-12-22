# spo report siteusagepages

Gets the number of pages viewed across all sites

## Usage

```sh
m365 spo report siteusagepages [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7`, `D30`, `D90`, `D180`.

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in.

--8<-- "docs/cmd/_global.md"

## Examples

Gets the number of pages viewed across all sites for the last week

```sh
m365 spo report siteusagepages --period D7
```

Gets the number of pages viewed across all sites for the last week and exports the report data in the specified path in text format

```sh
m365 spo report siteusagepages --period D7 --output text > "siteusagepages.txt"
```

Gets the number of pages viewed across all sites for the last week and exports the report data in the specified path in json format

```sh
m365 spo report siteusagepages --period D7 --output json > "siteusagepages.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-26",
        "Site Type": "All",
        "Page View Count": "14",
        "Report Date": "2022-11-26",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Site Type,Page View Count,Report Date,Report Period
    2022-11-26,All,14,2022-11-26,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Site Type,Page View Count,Report Date,Report Period
    2022-11-26,All,14,2022-11-26,7
    ```
