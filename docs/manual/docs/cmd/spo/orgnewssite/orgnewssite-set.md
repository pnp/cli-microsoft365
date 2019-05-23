# spo orgnewssite set

Marks a site as one of multiple possible organizational news sites for the tenant

## Usage

```sh
spo orgnewssite set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|Absolute URL of the site to add
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To add a site from the list of organizational news sites, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`
If you are logged in to a different site and try to manage tenant properties,
you will get an error.

## Examples

Set a site as an organizational news site

```sh
spo orgnewssite set -u https://contoso.sharepoint.com/sites/site1
```
