# spo hubsite disconnect

Disconnects the specifies site collection from its hub site

## Usage

```sh
spo hubsite disconnect [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|URL of the site collection to disconnect from its hub site
`--confirm`|Don't prompt for confirming disconnecting from the hub site
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To disconnect a site collection from its hub site, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`. If you are connected to a different site and will try to manage tenant properties, you will get an error.

## Examples

Disconnect the site collection with URL _https://contoso.sharepoint.com/sites/sales_ from its hub site. Will prompt for confirmation before disconnecting from the hub site.

```sh
spo hubsite disconnect --url https://contoso.sharepoint.com/sites/sales
```

Disconnect the site collection with URL _https://contoso.sharepoint.com/sites/sales- from its hub site without prompting for confirmation

```sh
spo hubsite disconnect --url https://contoso.sharepoint.com/sites/sales --confirm
```

## More information

- SharePoint hub sites new in Office 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)