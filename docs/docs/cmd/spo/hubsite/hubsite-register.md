# spo hubsite register

Registers the specified site collection as a hub site

## Usage

```sh
m365 spo hubsite register [options]
```

## Options

`-u, --url <url>`
: URL of the site collection to register as a hub site

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

If the specified site collection is already registered as a hub site, you will get a `This site is already a HubSite.` error.

## Examples

Register the site collection with URL _https://contoso.sharepoint.com/sites/sales_ as a hub site

```sh
m365 spo hubsite register --url https://contoso.sharepoint.com/sites/sales
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)