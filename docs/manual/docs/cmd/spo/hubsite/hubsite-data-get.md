# spo hubsite data get

Gets the hub site data for the specified site

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

## Usage

```sh
spo hubsite data get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --webUrl <webUrl>`|Web site Url
`-f, --forceRefresh`|If set, the cache is refreshed with latest updates

!!! important
    Before using this command, connect to a SharePoint Online site, using the [spo connect](../connect.md) command.

## Remarks

To get hub site data for a web, you have to first connect to a SharePoint site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso.sharepoint.com`.

If the specified `webUrl` is not connected to a hub site site, you will get a `odata.null: true` message.

## Examples

Get information about the hub site data for a web with Url https://contoso.sharepoint.com/sites/project-x

```sh
spo hubsite data get --webUrl https://contoso.sharepoint.com/sites/project-x
```

## More information

- SharePoint hub sites new in Office 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)