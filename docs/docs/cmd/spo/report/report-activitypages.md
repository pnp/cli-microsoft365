# spo report activitypages

Gets the number of unique pages visited by users

## Usage

```sh
m365 spo report activitypages [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7`, `D30`, `D90`, `D180`.

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in.

--8<-- "docs/cmd/_global.md"

## Examples

Gets the number of unique pages visited by users for the last week

```sh
m365 spo report activitypages --period D7
```

Gets the number of unique pages visited by users for the last week and exports the report data in the specified path in text format

```sh
m365 spo report activitypages --period D7 --output text > "activitypages.txt"
```

Gets the number of unique pages visited by users for the last week and exports the report data in the specified path in json format

```sh
m365 spo report activitypages --period D7 --output json > "activitypages.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-11-26",
        "Visited Page Count": "10",
        "Report Date": "2022-11-26",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Visited Page Count,Report Date,Report Period
    2022-11-26,10,2022-11-26,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Visited Page Count,Report Date,Report Period
    2022-11-26,10,2022-11-26,7
    ```
