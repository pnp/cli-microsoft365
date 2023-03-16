# skype report activitycounts

Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions

## Usage

```sh
m365 skype report activitycounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week

```sh
m365 skype report activitycounts --period D7
```

Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week and exports the report data in the specified path in text format

```sh
m365 skype report activitycounts --period D7 --output text > "activitycounts.txt"
```

Gets the trends on how many users organized and participated in conference sessions held in your organization through Skype for Business. The report also includes the number of peer-to-peer sessions for the last week and exports the report data in the specified path in json format

```sh
m365 skype report activitycounts --period D7 --output json > "activitycounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2023-01-25",
        "Report Date": "2023-01-19",
        "Report Period": "7",
        "Peer-to-peer": "0",
        "Organized": "0",
        "Participated": "0"
      }
    ]
    ```

=== "Text"

    ```txt
    Report Refresh Date,Report Date,Report Period,Peer-to-peer,Organized,Participated
    2023-01-25,2023-01-19,7,0,0,0
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Report Date,Report Period,Peer-to-peer,Organized,Participated
    2023-01-25,2023-01-19,7,0,0,0
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Report Date,Report Period,Peer-to-peer,Organized,Participated
    2023-01-25,2023-01-19,7,0,0,0
    ```
