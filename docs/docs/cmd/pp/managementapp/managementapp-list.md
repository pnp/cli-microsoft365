# pp managementapp list

Lists management applications for Power Platform

## Usage

```sh
m365 pp managementapp list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Examples

Lists management applications for Power Platform

```sh
m365 pp managementapp list
```

## Response

=== "JSON"

    ```json
    [
      {
        "applicationId":"31359c7f-bd7e-475c-86db-fdb8c937548e"
      }
    ]
    ```

=== "Text"

    ```text
    applicationId
    ------------------------------------
    31359c7f-bd7e-475c-86db-fdb8c937548e
    ```

=== "CSV"

    ```csv
    applicationId
    31359c7f-bd7e-475c-86db-fdb8c937548e
    ```

=== "Markdown"

    ```md
    # pp managementapp list

    Date: 9/1/2023

    ## undefined (undefined)

    Property | Value
    ---------|-------
    applicationId | 31359c7f-bd7e-475c-86db-fdb8c937548e
    ```
