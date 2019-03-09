# spo site inplacerecordsmanagement set

Activates or deactivates in-place records management for a site collection

## Usage

```sh
spo site inplacerecordsmanagement set [options]
```

## Options

Option|Description
------|-----------
`--help`|output usage information
`-u, --siteUrl <siteUrl>`|The URL of the site on which to activate or deactivate in-place records management
`--enabled <enabled>`|Set to `true` to activate in-place records management and to `false` to deactivate it
`-o, --output [output]`|Output type. `json|text`. Default `text`
`--verbose`|Runs command with verbose logging
`--debug`|Runs command with debug logging

!!! important
    Before using this command, log in to a SharePoint Online site, using the [spo login](../login.md) command.

## Remarks

To activate or deactivate in-place records management, you have to first log in to SharePoint using the [spo login](../login.md) command, eg. `spo login https://contoso.sharepoint.com`.

## Examples

Activates in-place records management for site _https://contoso.sharepoint.com/sites/team-a_

```sh
spo site inplacerecordsmanagement set --siteUrl https://contoso.sharepoint.com/sites/team-a --enabled true
```

Deactivates in-place records management for site _https://contoso.sharepoint.com/sites/team-a_

```sh
spo site inplacerecordsmanagement set --siteUrl https://contoso.sharepoint.com/sites/team-a --enabled false
```