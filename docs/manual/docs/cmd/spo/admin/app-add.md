# spo tenant appcatalog get

Get URL of the tenant app catalog

## Usage

```sh
spo tenant appcatalog get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To view the status of an Office 365 CDN, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`. If you are connected to a different site and will try to manage tenant properties, you will get an error.

## Examples

Get URL of the tenant app catalog

```sh
spo tenant appcatalog get
```