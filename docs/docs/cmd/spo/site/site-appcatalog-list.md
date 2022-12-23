# spo site appcatalog list

List all site collection app catalogs within the tenant

## Usage

```sh
m365 spo site appcatalog list [options]
```

## Options

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you have to have at least read permissions on the SharePoint root site.

## Examples

List all site collection app catalogs within the tenant

```sh
m365 spo site appcatalog list
```

## Response

=== "JSON"

    ```json
    [
      {
        "AbsoluteUrl": "https://contoso.sharepoint.com/sites/site1",
        "ErrorMessage": "Success",
        "SiteID": "9798e615-b586-455e-8486-84913f492c49"
      }
    ]
    ```

=== "Text"

    ```text
    AbsoluteUrl                                          SiteID
    ---------------------------------------------------  ------------------------------------
    https://contoso.sharepoint.com/sites/site1           9798e615-b586-455e-8486-84913f492c49
    ```

=== "CSV"

    ```csv
    AbsoluteUrl,SiteID
    https://contoso.sharepoint.com/sites/site1,9798e615-b586-455e-8486-84913f492c49
    ```
