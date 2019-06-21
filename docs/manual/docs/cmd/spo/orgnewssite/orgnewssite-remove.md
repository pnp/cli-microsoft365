# spo orgnewssite remove

Removes a site from the list of organizational news sites

## Usage

```sh
spo orgnewssite remove [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --url <url>`|Absolute URL of the site to remove
`--confirm`|Don't prompt for confirmation
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, connect to a SharePoint Online tenant admin site, using the [spo login](../login.md) command.

## Remarks

To remove a site from the list of organizational news sites, you have to first log in to a tenant admin site using the [spo login](../login.md) command, eg. `spo login https://contoso-admin.sharepoint.com`.

## Examples

Remove a site from the list of organizational news

```sh
spo orgnewssite remove --url https://contoso.sharepoint.com/sites/site1
```

Remove a site from the list of organizational news sites, without prompting for confirmation

```sh
spo orgnewssite remove --url https://contoso.sharepoint.com/sites/site1 --confirm
```