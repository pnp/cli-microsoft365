# outlook report mailboxusagequotastatusmailboxcounts

Gets the count of user mailboxes in each quota category

## Usage

```sh
m365 outlook report mailboxusagequotastatusmailboxcounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the count of user mailboxes in each quota category for the last week

```sh
m365 outlook report mailboxusagequotastatusmailboxcounts --period D7
```

Gets the count of user mailboxes in each quota category for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailboxusagequotastatusmailboxcounts --period D7 --output text > "mailboxusagequotastatusmailboxcounts.txt"
```

Gets the count of user mailboxes in each quota category for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailboxusagequotastatusmailboxcounts --period D7 --output json > "mailboxusagequotastatusmailboxcounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2023-01-24",
        "Under Limit": "146",
        "Warning Issued": "0",
        "Send Prohibited": "0",
        "Send/Receive Prohibited": "0",
        "Indeterminate": "0",
        "Report Date": "2023-01-18",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```txt
    Report Refresh Date,Under Limit,Warning Issued,Send Prohibited,Send/Receive Prohibited,Indeterminate,Report Date,Report Period
    2023-01-24,146,0,0,0,0,2023-01-18,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Under Limit,Warning Issued,Send Prohibited,Send/Receive Prohibited,Indeterminate,Report Date,Report Period
    2023-01-24,146,0,0,0,0,2023-01-18,7
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Under Limit,Warning Issued,Send Prohibited,Send/Receive Prohibited,Indeterminate,Report Date,Report Period
    2023-01-24,146,0,0,0,0,2023-01-18,7
    ```
