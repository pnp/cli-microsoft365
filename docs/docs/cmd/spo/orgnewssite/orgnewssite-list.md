# spo orgnewssite list

Lists all organizational news sites

## Usage

```sh
m365 spo orgnewssite list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

List all organizational news sites

```sh
m365 spo orgnewssite list
```

## Response

=== "JSON"

    ```json
    [
      "https://contoso.sharepoint.com/sites/contosoNews"
    ]
    ```

=== "Text"

    ```text
    https://contoso.sharepoint.com/sites/contosoNews
    ```

=== "CSV"

    ```csv
    https://contoso.sharepoint.com/sites/contosoNews
    ```

=== "Markdown"

    ```md
    https://contoso.sharepoint.com/sites/contosoNews
    ```
