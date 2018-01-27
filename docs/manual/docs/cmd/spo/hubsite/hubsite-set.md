# spo hubsite set

Updates properties of the specified hub site

!!! attention
    This command is based on a SharePoint API that is currently in preview and is subject to change once the API reached general availability.

## Usage

```sh
spo hubsite set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-i, --id <id>`|ID of the hub site to update
`-t, --title [title]`|The new title for the hub site
`-d, --description [description]`|The new description for the hub site
`-l, --logoUrl [logoUrl]`|The URL of the new logo for the hub site
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To update hub site's properties, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

If the specified `id` doesn't refer to an existing hub site, you will get an `Unknown Error` error.

## Examples

Update hub site's title

```sh
spo hubsite set --id 255a50b2-527f-4413-8485-57f4c17a24d1 --title Sales
```

Update hub site's title and description

```sh
spo hubsite set --id 255a50b2-527f-4413-8485-57f4c17a24d1 --title Sales --description "All things sales"
```

## More information

- SharePoint hub sites new in Office 365: [https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547](https://techcommunity.microsoft.com/t5/SharePoint-Blog/SharePoint-hub-sites-new-in-Office-365/ba-p/109547)