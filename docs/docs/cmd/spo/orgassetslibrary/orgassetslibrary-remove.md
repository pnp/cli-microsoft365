# spo orgassetslibrary remove

Removes a library that was designated as a central location for organization assets across the tenant.

## Usage

```sh
m365 spo orgassetslibrary remove [options]
```

## Options

`--libraryUrl <libraryUrl>`
: The server relative URL of the library to be removed as a central location for organization assets.

`--confirm`
: Don't prompt for confirming removing the organization asset library.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Removes organization assets library without prompting for confirmation

```sh
m365 spo orgassetslibrary remove --libraryUrl "/sites/branding/assets" --confirm
```

## Response

=== "JSON"

    ```json
    {
      "IsNull": false
    }
    ```

=== "Text"

    ```text
    IsNull: false
    ```

=== "CSV"

    ```csv
    IsNull
    ```

=== "Markdown"

    ```md
    # spo orgassetslibrary remove --libraryUrl "https://contoso.sharepoint.com/sites/branding/SiteAssets" --confirm "true"

    Date: 5/1/2023

    Property | Value
    ---------|-------
    IsNull | false
    ```
