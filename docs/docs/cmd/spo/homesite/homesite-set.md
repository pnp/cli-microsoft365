# spo homesite set

Sets the specified site as the Home Site

## Usage

```sh
m365 spo homesite set [options]
```

## Options

`-u, --siteUrl <siteUrl>`
: The URL of the site to set as Home Site

`-v, --VivaConnectionsDefaultStart [VivaConnectionsDefaultStart]`
: When set to $true, the VivaConnectionsDefaultStart parameter will keep the Viva Connections landing experience to the SharePoint home site. If set to $false, the Viva Connections home experience will be used. 

--8<-- "docs/cmd/_global.md"

!!! important
    To use this command you have to have permissions to access the tenant admin site.

## Examples
### EXAMPLE 1
Set the specified site as the Home Site

```sh
m365 spo homesite set --siteUrl https://contoso.sharepoint.com/sites/comms
```
## Response

=== "JSON"

    ```json
    "The Home site has been set to https://contoso.sharepoint.com. It may take some time for the change to apply. Check aka.ms/homesites for details."
    ```

=== "Text"

    ```text
    The Home site has been set to https://contoso.sharepoint.com. It may take some time for the change to apply. Check aka.ms/homesites for details.
    ```

=== "CSV"

    ```csv
    The Home site has been set to https://contoso.sharepoint.com. It may take some time for the change to apply. Check aka.ms/homesites for details.
    ```
### EXAMPLE 2
Sets the home site to the provided site collection url and keeps the Viva Connections landing experience to the SharePoint home site.

```sh
m365 spo homesite set --siteUrl https://contoso.sharepoint.com/sites/comms --VivaConnectionsDefaultStart $true
```

=== "JSON"

    ```json
    "The Home site has been set to https://contoso.sharepoint.com and the Viva Connections default experience to True. It may take some time for the change to apply. Check aka.ms/homesites for details."
    ```

=== "Text"

    ```text
    The Home site has been set to https://contoso.sharepoint.com and the Viva Connections default experience to True. It may take some time for the change to apply. Check aka.ms/homesites for details.
    ```

=== "CSV"

    ```csv
    The Home site has been set to https://contoso.sharepoint.com and the Viva Connections default experience to True. It may take some time for the change to apply. Check aka.ms/homesites for details.
    ```

## More information

- SharePoint home sites: a landing for your organization on the intelligent intranet: [https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933](https://techcommunity.microsoft.com/t5/Microsoft-SharePoint-Blog/SharePoint-home-sites-a-landing-for-your-organization-on-the/ba-p/621933)
