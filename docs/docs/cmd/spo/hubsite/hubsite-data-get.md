# spo hubsite data get

Get hub site data for the specified site

## Usage

```sh
m365 spo hubsite data get [options]
```

## Options

`-h, --help`
: output usage information

`-u, --webUrl <webUrl>`
: URL of the site for which to retrieve hub site data

`-f, --forceRefresh`
: Set, to refresh the server cache with the latest updates

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

By default, the hub site data is returned from the server's cache. To refresh the data with the latest updates, use the `-f, --forceRefresh` option. Use this option, if you just made changes and need to see them right
away.

If the specified site is not connected to a hub site site and is not a hub site itself, no data will be retrieved.

## Examples

Get information about the hub site data for a site with URL https://contoso.sharepoint.com/sites/project-x

```sh
m365 spo hubsite data get --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- SharePoint hub sites new in Microsoft 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)