# spo orgnewssite set

Marks site as an organizational news site

## Usage

```sh
spo orgnewssite set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|The URL of the site to mark as an organizational news site
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To add a site from the list of organizational news sites, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

Using the `-u, --url` option you can specify which site to add to the list of organizational news sites.

## Examples

Set a site as an organizational news site

```sh
spo orgnewssite set --url https://contoso.sharepoint.com/sites/site1
```
