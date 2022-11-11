# yammer report activityusercounts

Gets the trends on the number of unique users who posted, read, and liked Yammer messages

## Usage

```sh
m365 yammer report activityusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trends on the number of unique users who posted, read, and liked Yammer messages for the last week

```sh
m365 yammer report activityusercounts --period D7
```

Gets the trends on the number of unique users who posted, read, and liked Yammer messages for the last week and exports the report data in the specified path in text format

```sh
m365 yammer report activityusercounts --period D7 --output text > "activityusercounts.txt"
```

Gets the trends on the number of unique users who posted, read, and liked Yammer messages for the last week and exports the report data in the specified path in json format

```sh
m365 yammer report activityusercounts --period D7 --output json > "activityusercounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-09",
        "Liked": "9",
        "Posted": "5",
        "Read": "64",
        "Report Date": "2022-11-09",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Liked,Posted,Read,Report Date,Report Period
    2022-11-09,9,5,64,2022-11-09,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Liked,Posted,Read,Report Date,Report Period
    2022-11-09,9,5,64,2022-11-09,7
    ```
