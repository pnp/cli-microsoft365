# yammer report activitycounts

Gets the trends on the amount of Yammer activity in your organization by how many messages were posted, read, and liked

## Usage

```sh
m365 yammer report activitycounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trends on the amount of Yammer activity in your organization by how many messages were posted, read, and liked for the last week

```sh
m365 yammer report activitycounts --period D7
```

Gets the trends on the amount of Yammer activity in your organization by how many messages were posted, read, and liked for the last week and exports the report data in the specified path in text format

```sh
m365 yammer report activitycounts --period D7 --output text > "activitycounts.txt"
```

Gets the trends on the amount of Yammer activity in your organization by how many messages were posted, read, and liked for the last week and exports the report data in the specified path in json format

```sh
m365 yammer report activitycounts --period D7 --output json > "activitycounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-09",
        "Liked": "12",
        "Posted": "6",
        "Read": "435",
        "Report Date": "2022-11-09",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Liked,Posted,Read,Report Date,Report Period
    2022-11-09,12,6,435,2022-11-09,90
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Liked,Posted,Read,Report Date,Report Period
    2022-11-09,12,6,435,2022-11-09,90
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Liked,Posted,Read,Report Date,Report Period
    2022-11-09,12,6,435,2022-11-09,90
    ```
