# outlook report mailboxusagestorage

Gets the amount of mailbox storage used in your organization

## Usage

```sh
m365 outlook report mailboxusagestorage [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the amount of mailbox storage used in your organization for the last week

```sh
m365 outlook report mailboxusagestorage --period D7
```

Gets the amount of mailbox storage used in your organization for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailboxusagestorage --period D7 --output text > "mailboxusagestorage.txt"
```

Gets the amount of mailbox storage used in your organization for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailboxusagestorage --period D7 --output json > "mailboxusagestorage.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2023-01-24",
        "Storage Used (Byte)": "791481624264",
        "Report Date": "2023-01-18",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```txt
    Report Refresh Date,Storage Used (Byte),Report Date,Report Period
    2023-01-24,791481624264,2023-01-18,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Storage Used (Byte),Report Date,Report Period
    2023-01-24,791481624264,2023-01-18,7
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Storage Used (Byte),Report Date,Report Period
    2023-01-24,791481624264,2023-01-18,7
    ```
