# spo site get

Gets information about the specific site collection

## Usage

```sh
spo site get [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|URL of the site to retrieve information for
`-o, --output <output>`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo connect](../connect.md) command.

## Remarks

To get information about a site collection, you have to first connect to a tenant admin site using the [spo connect](../connect.md) command, eg. `spo connect https://contoso-admin.sharepoint.com`. If you are connected to a different site and will try to to get site collection information, you will get an error.

This command can retrieve information for both classic and modern sites.

## Examples

Return information about the _https://contoso.sharepoint.com/sites/project-x_ site collection.

```sh
spo site get -u https://contoso.sharepoint.com/sites/project-x
```
