# spo hubsite unregister

Unregisters the specifies site collection as a hub site

## Usage

```sh
m365 spo hubsite unregister [options]
```

## Options

`-h, --help`
: output usage information

`-u, --url <url>`
: URL of the site collection to unregister as a hub site

`--confirm`
: Don't prompt for confirming unregistering the hub site

`--query [query]`
: JMESPath query string. See [http://jmespath.org/](http://jmespath.org/) for more information and examples

`-o, --output [output]`
: Output type. `json,text`. Default `text`

`--verbose`
: Runs command with verbose logging

`--debug`
: Runs command with debug logging

## Remarks

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

If the specified site collection is not registered as a hub site, you will get a `hubSiteId` error.

## Examples

Unregister the site collection with URL _https://contoso.sharepoint.com/sites/sales_ as a hub site. Will prompt for confirmation before unregistering the hub site.

```sh
m365 spo hubsite unregister --url https://contoso.sharepoint.com/sites/sales
```

Unregister the site collection with URL _https://contoso.sharepoint.com/sites/sales_ as a hub site without prompting for confirmation

```sh
m365 spo hubsite unregister --url https://contoso.sharepoint.com/sites/sales --confirm
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)