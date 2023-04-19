# outlook report mailboxusagemailboxcount

Gets the total number of user mailboxes in your organization and how many are active each day of the reporting period.

## Usage

```sh
m365 outlook report mailboxusagemailboxcount [options]
```

## Options

`-p, --period <period>`
: The length of time over which the report is aggregated. Supported values `D7,D30,D90,D180`

`-f, --outputFile [outputFile]`
: Path to the file where the report should be stored in

--8<-- "docs/cmd/_global.md"

## Remarks

A mailbox is considered active if the user sent or read any email.

## Examples

Gets the total number of user mailboxes in your organization and how many are active each day for the last week.

```sh
m365 outlook report mailboxusagemailboxcount --period D7
```

Gets the total number of user mailboxes in your organization and how many are active each day for the last week and exports the report data in the specified path in text format

```sh
m365 outlook report mailboxusagemailboxcount --period D7 --output text > "mailboxusagemailboxcount.txt"
```

Gets the total number of user mailboxes in your organization and how many are active each day for the last week and exports the report data in the specified path in json format

```sh
m365 outlook report mailboxusagemailboxcount --period D7 --output json > "mailboxusagemailboxcount.json"
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2023-01-24",
        "Total": "146",
        "Active": "131",
        "Report Date": "2023-01-18",
        "Report Period": "7"
      }
    ]
    ```

=== "Text"

    ```txt
    Report Refresh Date,Total,Active,Report Date,Report Period
    2023-01-24,146,131,2023-01-18,7
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Total,Active,Report Date,Report Period
    2023-01-24,146,131,2023-01-18,7
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Total,Active,Report Date,Report Period
    2023-01-24,146,131,2023-01-18,7
    ```
