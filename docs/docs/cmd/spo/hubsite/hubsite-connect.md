# spo hubsite connect

Connects the specified site collection to the given hub site

## Usage

```sh
m365 spo hubsite connect [options]
```

## Options

`-u, --url <url>`
: The URL of the site collection to connect to the hub site

`-i, --hubSiteId <hubSiteId>`
: The ID of the hub site to which to connect the site collection

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

If the specified site collection is already connected to a hub site, it will be disconnected and connected to the newly specified hub site.

If the specified `hubSiteId` doesn't point to a valid hub site, you will get a `ResourceNotFoundException` error.

## Examples

Connect the site collection with URL _https://contoso.sharepoint.com/sites/contoso-sales_ to the hub site with ID _255a50b2-527f-4413-8485-57f4c17a24d1_

```sh
m365 spo hubsite connect --url https://contoso.sharepoint.com/sites/contoso-sales --hubSiteId 255a50b2-527f-4413-8485-57f4c17a24d1
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)