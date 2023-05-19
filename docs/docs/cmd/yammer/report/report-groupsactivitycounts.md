# yammer report groupsactivitycounts

Gets the number of Yammer messages posted, read, and liked in groups

## Usage

```sh
m365 yammer report groupsactivitycounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the number of Yammer messages posted, read, and liked in groups for the last week

```sh
m365 yammer report groupsactivitycounts --period D7
```

Gets the number of Yammer messages posted, read, and liked in groups for the last week and exports the report data in the specified path in text format

```sh
m365 yammer report groupsactivitycounts --period D7 --output text > "groupsactivitycounts.txt"
```

Gets the number of Yammer messages posted, read, and liked in groups for the last week and exports the report data in the specified path in json format

```sh
m365 yammer report groupsactivitycounts --period D7 --output json > "groupsactivitycounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-04",
        "Liked": "5",
        "Posted": "6",
        "Read": "7",
        "Report Date": "2022-11-04",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Liked,Posted,Read,Report Date,Report Period
    2022-11-10,5,6,7,2022-11-10,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Liked,Posted,Read,Report Date,Report Period
    2022-11-10,5,6,7,2022-11-10,7
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Liked,Posted,Read,Report Date,Report Period
    2022-11-10,5,6,7,2022-11-10,7
    ```
