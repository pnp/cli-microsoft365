# spo hubsite connect

Connects the specified site collection to the given hub site

## Usage

```sh
spo hubsite connect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|The URL of the site collection to connect to the hub site
`-i, --hubSiteId <hubSiteId>`|The ID of the hub site to which to connect the site collection
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To connect a site collection to a hub site, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified site collection is already connected to a hub site, it will be disconnected and connected to the newly specified hub site.

If the specified `hubSiteId` doesn't point to a valid hub site, you will get a `ResourceNotFoundException` error.

## Examples

Connect the site collection with URL _https://contoso.sharepoint.com/sites/contoso-sales_ to the hub site with ID _255a50b2-527f-4413-8485-57f4c17a24d1_

```sh
spo hubsite connect --url https://contoso.sharepoint.com/sites/contoso-sales --hubSiteId 255a50b2-527f-4413-8485-57f4c17a24d1
```

## More information

- SharePoint hub sites new in Office 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)