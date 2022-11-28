# spo report siteusagefilecounts

Get the total number of files across all sites and the number of active files

## Usage

```sh
m365 spo report siteusagefilecounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7`, `D30`, `D90`, `D180`.

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in.

--8<-- "docs/cmd/_global.md"

## Remarks

A file (user or system) is considered active if it has been saved, synced, modified, or shared within the specified time period.

## Examples

Get the total number of files across all sites and the number of active files for the last week

```sh
m365 spo report siteusagefilecounts --period D7
```

Get the total number of files across all sites and the number of active files for the last week and exports the report data in the specified path in text format

```sh
m365 spo report siteusagefilecounts --period D7 --output text > "siteusagefilecounts.txt"
```

Get the total number of files across all sites and the number of active files for the last week and exports the report data in the specified path in json format

```sh
m365 spo report siteusagefilecounts --period D7 --output json > "siteusagefilecounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-26",
        "Site Type": "All",
        "Total": "1320",
        "Active": "3",
        "Report Date": "2022-11-26",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
    2022-11-26,All,1320,3,2022-11-26,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
    2022-11-26,All,1320,3,2022-11-26,7
    ```
