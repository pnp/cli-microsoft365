# onedrive report usagefilecounts

Gets the total number of files across all sites and how many are active files

## Usage

```sh
m365 onedrive report usagefilecounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

A file is considered active if it has been saved, synced, modified, or shared within the specified time period.

## Examples

Gets the total number of files across all sites and how many are active files for the last week

```sh
m365 onedrive report usagefilecounts --period D7
```

Gets the total number of files across all sites and how many are active files for the last week and exports the report data in the specified path in text format

```sh
m365 onedrive report usagefilecounts --period D7 --output text > "usagefilecounts.txt"
```

Gets the total number of files across all sites and how many are active files for the last week and exports the report data in the specified path in json format

```sh
m365 onedrive report usagefilecounts --period D7 --output json > "usagefilecounts.json"
```

## Response

=== "JSON"

```json
[
  {
    "Report Refresh Date": "2022-10-25",
    "Site Type": "All",
    "Total": "581190",
    "Active": "88",
    "Report Date": "2022-10-25",
    "Report Period": "7"
  },
  {
    "Report Refresh Date": "2022-10-25",
    "Site Type": "All",
    "Total": "581190",
    "Active": "394",
    "Report Date": "2022-10-24",
    "Report Period": "7"
  }
]
```

=== "Text"

    ``` text

Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
2022-10-25,All,581190,88,2022-10-25,7
2022-10-25,All,581190,394,2022-10-24,7
2022-10-25,All,581190,95,2022-10-23,7
2022-10-25,All,581096,19,2022-10-22,7
2022-10-25,All,581051,87,2022-10-21,7
2022-10-25,All,581051,61,2022-10-20,7
2022-10-25,All,580954,251,2022-10-19,7

````

=== "CSV"

    ``` text
Report Refresh Date,Site Type,Total,Active,Report Date,Report Period
2022-10-25,All,581190,88,2022-10-25,7
2022-10-25,All,581190,394,2022-10-24,7
2022-10-25,All,581190,95,2022-10-23,7
2022-10-25,All,581096,19,2022-10-22,7
2022-10-25,All,581051,87,2022-10-21,7
2022-10-25,All,581051,61,2022-10-20,7
2022-10-25,All,580954,251,2022-10-19,7

````
