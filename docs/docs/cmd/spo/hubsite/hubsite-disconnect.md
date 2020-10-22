# spo hubsite disconnect

Disconnects the specifies site collection from its hub site

## Usage

```sh
m365 spo hubsite disconnect [options]
```

## Options

`-u, --url <url>`
: URL of the site collection to disconnect from its hub site

`--confirm`
: Don't prompt for confirming disconnecting from the hub site

--8<-- "docs/cmd/_global.md"

## Remarks

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

## Examples

Disconnect the site collection with URL _https://contoso.sharepoint.com/sites/sales_ from its hub site. Will prompt for confirmation before disconnecting from the hub site.

```sh
m365 spo hubsite disconnect --url https://contoso.sharepoint.com/sites/sales
```

Disconnect the site collection with URL _https://contoso.sharepoint.com/sites/sales- from its hub site without prompting for confirmation

```sh
m365 spo hubsite disconnect --url https://contoso.sharepoint.com/sites/sales --confirm
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)