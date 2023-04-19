# outlook report mailappusageversionsusercounts

Gets the count of unique users by Outlook desktop version.

## Usage

```sh
m365 outlook report mailappusageversionsusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the count of unique users by Outlook desktop version for the last week

```sh
m365 outlook report mailappusageversionsusercounts --period D7
```

Gets the count of unique users by Outlook desktop version for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailappusageversionsusercounts --period D7 --output text > "mailappusageversionsusercounts.txt"
```

Gets the count of unique users by Outlook desktop version for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailappusageversionsusercounts --period D7 --output json > "mailappusageversionsusercounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2023-01-25",
        "Outlook 2016": "",
        "Outlook 2013": "",
        "Outlook 2010": "",
        "Outlook 2007": "",
        "Undetermined": "",
        "Report Period": "7",
        "Outlook M365": "108",
        "Outlook 2019": ""
      }
    ]
    ```

=== "Text"

    ```txt
    Report Refresh Date,Outlook 2016,Outlook 2013,Outlook 2010,Outlook 2007,Undetermined,Report Period,Outlook M365,Outlook 2019
    2023-01-25,,,,,,7,108,
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Outlook 2016,Outlook 2013,Outlook 2010,Outlook 2007,Undetermined,Report Period,Outlook M365,Outlook 2019
    2023-01-25,,,,,,7,108,
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Outlook 2016,Outlook 2013,Outlook 2010,Outlook 2007,Undetermined,Report Period,Outlook M365,Outlook 2019
    2023-01-25,,,,,,7,108,
    ```
