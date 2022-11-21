# spo homesite remove

Removes the current Home Site

## Usage

```sh
m365 spo homesite remove [options]
```

## Options

`--confirm`
: Do not prompt for confirmation before removing the Home Site

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples

Removes the current Home Site without confirmation

```sh
m365 spo homesite remove --confirm
```

## Response

=== "JSON"

    ```json
    "https://contoso.sharepoint.com has been removed as a Home site. It may take some time for the change to apply. Check aka.ms/homesites for details."
    ```

=== "Text"

    ```text
    https://contoso.sharepoint.com has been removed as a Home site. It may take some time for the change to apply. Check aka.ms/homesites for details.
    ```

=== "CSV"

    ```csv
    https://contoso.sharepoint.com has been removed as a Home site. It may take some time for the change to apply. Check aka.ms/homesites for details.
    ```

## More information

- SharePoint home sites, a landing for your organization on the intelligent intranet: [https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933](https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933)
