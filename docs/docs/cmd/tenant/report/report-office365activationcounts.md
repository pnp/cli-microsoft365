# tenant report office365activationcounts

Get the count of Microsoft 365 activations on desktops and devices

## Usage

```sh
m365 tenant report office365activationcounts [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Get the count of Microsoft 365 activations on desktops and devices

```sh
m365 tenant report office365activationcounts
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-10-25",
        "Product Type": "MICROSOFT 365 APPS FOR ENTERPRISE",
        "Windows": 5,
        "Mac": 0,
        "Android": 0,
        "iOS": 0,
        "Windows 10 Mobile": 0
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,Product Type,Windows,Mac,Android,iOS,Windows 10 Mobile
    2022-10-25,MICROSOFT 365 APPS FOR ENTERPRISE,5,0,0,0,0
    ```

=== "CSV"

    ```csv
    Report Refresh Date,Product Type,Windows,Mac,Android,iOS,Windows 10 Mobile
    2022-10-25,MICROSOFT 365 APPS FOR ENTERPRISE,5,0,0,0,0
    ```

=== "Markdown"

    ```md
    Report Refresh Date,Product Type,Windows,Mac,Android,iOS,Windows 10 Mobile
    2022-10-25,MICROSOFT 365 APPS FOR ENTERPRISE,5,0,0,0,0
    ```
