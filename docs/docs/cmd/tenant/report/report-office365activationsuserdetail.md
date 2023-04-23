# tenant report office365activationsuserdetail

Get details about users who have activated Microsoft 365

## Usage

```sh
m365 tenant report office365activationsuserdetail [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Get details about users who have activated Microsoft 365

```sh
m365 tenant report office365activationsuserdetail
```

## Response

=== "JSON"

    ```json
    [
      {
        "Report Refresh Date": "2022-10-25",
        "User Principal Name": "554630B7593DDE8E04F27933A965D5B2",
        "Display Name": "2BA57CC6348F53C9EE6347528301E896",
        "Product Type": "MICROSOFT EXCEL ADVANCED ANALYTICS",
        "Last Activated Date": "2021-01-10",
        "Windows": 0,
        "Mac": 0,
        "Windows 10 Mobile": 0,
        "iOS": 0,
        "Android": 0,
        "Activated On Shared Computer": "False"
      }
    ]
    ```

=== "Text"

    ```text
    Report Refresh Date,User Principal Name,Display Name,Product Type,Last Activated Date,Windows,Mac,Windows 10 Mobile,iOS,Android,Activated On Shared Computer
    2022-10-25,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,MICROSOFT 365 APPS FOR ENTERPRISE,2021-01-10,1,0,0,0,0,False
    ```

=== "CSV"

    ```csv
    Report Refresh Date,User Principal Name,Display Name,Product Type,Last Activated Date,Windows,Mac,Windows 10 Mobile,iOS,Android,Activated On Shared Computer
    2022-10-25,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,MICROSOFT 365 APPS FOR ENTERPRISE,2021-01-10,1,0,0,0,0,False
    ```

=== "Markdown"

    ```md
    Report Refresh Date,User Principal Name,Display Name,Product Type,Last Activated Date,Windows,Mac,Windows 10 Mobile,iOS,Android,Activated On Shared Computer
    2022-10-25,77E5979DD60BA6EAA53E814DBEEEFA5F,4291DA7C39EE3263E97336B42734A667,MICROSOFT 365 APPS FOR ENTERPRISE,2021-01-10,1,0,0,0,0,False
    ```
