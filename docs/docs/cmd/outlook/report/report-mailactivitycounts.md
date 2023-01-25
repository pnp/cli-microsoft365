# outlook report mailactivitycounts

Enables you to understand the trends of email activity (like how many were sent, read, and received) in your organization

## Usage

```sh
m365 outlook report mailactivitycounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week

```sh
m365 outlook report mailactivitycounts --period D7
```

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailactivitycounts --period D7 --output text > "mailactivitycounts.txt"
```

Gets the trends of email activity (like how many were sent, read, and received) in your organization for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailactivitycounts --period D7 --output json > "mailactivitycounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2023-01-24",
        "Send": "",
        "Receive": "2",
        "Read": "",
        "Meeting Created": "0",
        "Meeting Interacted": "",
        "Report Date": "2023-01-18",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```txt
    Report Refresh Date,Send,Receive,Read,Meeting Created,Meeting Interacted,Report Date,Report Period
    2023-01-24,,2,,0,,2023-01-18,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Send,Receive,Read,Meeting Created,Meeting Interacted,Report Date,Report Period
    2023-01-24,,2,,0,,2023-01-18,7
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Send,Receive,Read,Meeting Created,Meeting Interacted,Report Date,Report Period
    2023-01-24,,2,,0,,2023-01-18,7
    ```
