# spo knowledgehub set

Sets the Knowledge Hub Site for your tenant

## Usage

```sh
m365 spo knowledgehub set [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: URL of the site to set as Knowledge Hub

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Remarks

If the specified url doesn't refer to an existing site collection, you will get a `404 - "404 FILE NOT FOUND"` error.

## Examples

Sets the Knowledge Hub Site for your tenant

```sh
m365 spo knowledgehub set --siteUrl https://contoso.sharepoint.com/sites/knowledgesite
```

## Response

=== "JSON"

    ```json
    "The knowledge hub site with url \"https://contoso.sharepoint.com\" is added to list."
    ```

=== "Text"

    ```text
    The knowledge hub site with url "https://contoso.sharepoint.com" is added to list.
    ```

=== "CSV"

    ```csv
    The knowledge hub site with url "https://contoso.sharepoint.com" is added to list.
    ```
