# spo site set

Updates properties of the specified site

## Usage

```sh
spo site set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|The URL of the site collection to update
`-i, --id [id]`|The ID of the site collection to update
`--classification [classification]`|The new classification for the site collection
`--disableFlows [disableFlows]`|Set to `true` to disable using Microsoft Flow in this site collection
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To update site collection's properties, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`.

If the specified url doesn't refer to an existing site collection, you will get a `404 - "404 FILE NOT FOUND"` error.

To update site collection's properties, the command requires site collection ID. The command can retrieve it automatically, but if you already have it, you can save an additional request, by specifying it using the `id` option.

## Examples

Update site collection's classification. Will automatically retrieve the ID of the site collection

```sh
spo site set --url https://contoso.sharepoint.com/sites/sales --classification MBI
```

Reset site collection's classification.

```sh
spo site set --url https://contoso.sharepoint.com/sites/sales --id 255a50b2-527f-4413-8485-57f4c17a24d1 --classification
```

Disable using Microsoft Flow on the site collection. Will automatically retrieve the ID of the site collection

```sh
spo site set --url https://contoso.sharepoint.com/sites/sales --disableFlows true
```