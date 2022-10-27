# tenant report office365activationsusercounts

Get the count of users that are enabled and those that have activated the Office subscription on desktop or devices or shared computers

## Usage

```sh
m365 tenant report office365activationsusercounts [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Get the count of users that are enabled and those that have activated the Office subscription on desktop or devices or shared computers

```sh
m365 tenant report office365activationsusercounts
```

## Response

=== "JSON"

    ``` json
    [
      {
        "Report Refresh Date": "2022-10-25",
        "Product Type": "MICROSOFT 365 APPS FOR ENTERPRISE",
        "Assigned": 24,
        "Activated": 5,
        "Shared Computer Activation": 0
      }
    ]
    ```

=== "Text"

    ``` text
    Report Refresh Date,Product Type,Assigned,Activated,Shared Computer Activation
    2022-10-25,MICROSOFT 365 APPS FOR ENTERPRISE,24,5,0
    ```

=== "CSV"

    ``` CSV
    Report Refresh Date,Product Type,Assigned,Activated,Shared Computer Activation
    2022-10-25,MICROSOFT 365 APPS FOR ENTERPRISE,24,5,0
    ```
