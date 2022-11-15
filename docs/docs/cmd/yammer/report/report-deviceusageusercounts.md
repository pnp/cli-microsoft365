# yammer report deviceusageusercounts

Gets the number of daily users by device type

## Usage

```sh
m365 yammer report deviceusageusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the number of daily users by device type for the last week

```sh
m365 yammer report deviceusageusercounts --period D7
```

Gets the number of daily users by device type for the last week and exports the report data in the specified path in text format

```sh
m365 yammer report deviceusageusercounts --period D7 --output text > "deviceusageusercounts.txt"
```

Gets the number of daily users by device type for the last week and exports the report data in the specified path in json format

```sh
m365 yammer report deviceusageusercounts --period D7 --output json > "deviceusageusercounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-09",
        "Web": "4",
        "Windows Phone": "5",
        "Android Phone": "6",
        "iPhone": "3",
        "iPad": "3",
        "Other": "60",
        "Report Date": "2022-11-09",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Web,Windows Phone,Android Phone,iPhone,iPad,Other,Report Date,Report Period
    2022-11-09,4,5,6,3,3,60,2022-11-09,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Web,Windows Phone,Android Phone,iPhone,iPad,Other,Report Date,Report Period
    2022-11-09,4,5,6,3,3,60,2022-11-09,7
    ```
