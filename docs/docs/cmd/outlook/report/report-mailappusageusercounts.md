# outlook report mailappusageusercounts

Gets the count of unique users that connected to Exchange Online using any email app

## Usage

```sh
m365 outlook report mailappusageusercounts [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Examples

Gets the count of unique users that connected to Exchange Online using any email app for the last week

```sh
m365 outlook report mailappusageusercounts --period D7
```

Gets the count of unique users that connected to Exchange Online using any email app for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailappusageusercounts --period D7 --output text > "mailappusageusercounts.txt"
```

Gets the count of unique users that connected to Exchange Online using any email app for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailappusageusercounts --period D7 --output json > "mailappusageusercounts.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2023-01-25",
        "Mail For Mac": "",
        "Outlook For Mac": "2",
        "Outlook For Windows": "99",
        "Outlook For Mobile": "46",
        "Other For Mobile": "",
        "Outlook For Web": "",
        "POP3 App": "",
        "IMAP4 App": "",
        "SMTP App": "",
        "Report Date": "2023-01-19",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```txt
    Report Refresh Date,Mail For Mac,Outlook For Mac,Outlook For Windows,Outlook For Mobile,Other For Mobile,Outlook For Web,POP3 App,IMAP4 App,SMTP App,Report Date,Report Period
    2023-01-25,,2,99,46,,,,,,2023-01-19,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Mail For Mac,Outlook For Mac,Outlook For Windows,Outlook For Mobile,Other For Mobile,Outlook For Web,POP3 App,IMAP4 App,SMTP App,Report Date,Report Period
    2023-01-25,,2,99,46,,,,,,2023-01-19,7
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Mail For Mac,Outlook For Mac,Outlook For Windows,Outlook For Mobile,Other For Mobile,Outlook For Web,POP3 App,IMAP4 App,SMTP App,Report Date,Report Period
    2023-01-25,,2,99,46,,,,,,2023-01-19,7
    ```
